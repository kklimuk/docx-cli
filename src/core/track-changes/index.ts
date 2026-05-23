import type { DocView } from "../ast/doc-view";
import { XmlNode } from "../parser";

export type TrackedMeta = {
	author: string;
	date: string;
	revisionId: number;
};

export type RevisionAllocator = { next(): number };

export function isTrackChangesEnabled(view: DocView): boolean {
	if (!view.settingsTree) return false;
	const settingsRoot = XmlNode.findRoot(view.settingsTree, "w:settings");
	if (!settingsRoot) return false;
	return settingsRoot.children.some((child) => child.tag === "w:trackChanges");
}

export function createRevisionAllocator(view: DocView): RevisionAllocator {
	let nextId = computeMaxRevisionId(view) + 1;
	return {
		next(): number {
			const id = nextId;
			nextId += 1;
			return id;
		},
	};
}

export function resolveAuthor(authorFlag?: string): string {
	if (authorFlag) return authorFlag;
	if (Bun.env.DOCX_AUTHOR) return Bun.env.DOCX_AUTHOR;
	return "docx-cli";
}

export function resolveDate(): string {
	return Bun.env.DOCX_CLI_NOW ?? new Date().toISOString();
}

/** Mint a fresh `TrackedMeta` (author + date + revisionId) for a tracked
 * change about to be emitted. Convenience over the `resolveAuthor + resolveDate
 * + createRevisionAllocator(view).next()` triple that every mutating CLI verb
 * needs once per call. */
export function mintRevisionMeta(
	view: DocView,
	authorFlag?: string,
): TrackedMeta {
	return {
		author: resolveAuthor(authorFlag),
		date: resolveDate(),
		revisionId: createRevisionAllocator(view).next(),
	};
}

export function convertTextToDelText(node: XmlNode): XmlNode {
	const cloned = node.clone();
	mutateTextToDelText([cloned]);
	return cloned;
}

function computeMaxRevisionId(view: DocView): number {
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
	walkXml(view.documentTree, visit);
	if (view.footnotesTree) walkXml(view.footnotesTree, visit);
	if (view.endnotesTree) walkXml(view.endnotesTree, visit);
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
