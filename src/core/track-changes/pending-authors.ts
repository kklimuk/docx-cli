import type { XmlNode } from "../parser";
import { resolveAuthor } from "./index";

/** Watch `parent[startIndex..endIndex]` across an in-place content rebuild
 *  (`Edit.paragraph` / `Edit.range`) for other authors' pending revisions it
 *  resolves without review (issue #17). Call BEFORE mutating; the returned
 *  function, called after, names (in document order) every author other than
 *  the editor whose run-level revision (`<w:ins>`/`<w:del>`/`<w:moveFrom>`/
 *  `<w:moveTo>`) or `<w:pPrChange>` no longer appears in the edited region.
 *
 *  Before/after instead of predicting: each rebuild path keeps a different
 *  subset — the `--text` diff carries text boxes and inline content controls
 *  through and keeps the pPr; the `--markdown`/`--runs` replace drops the box
 *  and a pPrChange, and under tracking keeps an existing `<w:del>` but drops
 *  every `<w:ins>` — so only diffing what survived stays accurate as those
 *  paths change. A revision is matched by tag + `w:id` + author + date, so a
 *  CLONED survivor (the paragraph-mark marker an inherited pPr carries over)
 *  counts as kept. The region's extent after the edit is read off the parent's
 *  length change: a replace splices N paragraphs where M stood. */
export function watchPendingRevisions(
	parent: XmlNode[],
	startIndex: number,
	endIndex: number,
	authorFlag: string | undefined,
): () => string[] {
	const editor = resolveAuthor(authorFlag);
	const pending = new Map<string, string>();
	for (const node of parent.slice(startIndex, endIndex + 1)) {
		visitRevisions(node, (revision) => {
			const author = revision.getAttribute("w:author") ?? "";
			if (author === editor) return;
			pending.set(revisionKey(revision), author || "(unknown author)");
		});
	}
	const lengthBefore = parent.length;
	return () => {
		if (pending.size === 0) return [];
		const editedEnd = endIndex + 1 + parent.length - lengthBefore;
		const surviving = new Set<string>();
		for (const node of parent.slice(startIndex, editedEnd)) {
			visitRevisions(node, (revision) => surviving.add(revisionKey(revision)));
		}
		const authors = new Set<string>();
		for (const [key, author] of pending) {
			if (!surviving.has(key)) authors.add(author);
		}
		return [...authors];
	};
}

const REVISION_TAGS: ReadonlySet<string> = new Set([
	"w:ins",
	"w:del",
	"w:moveFrom",
	"w:moveTo",
	"w:pPrChange",
]);

function visitRevisions(
	node: XmlNode,
	visit: (revision: XmlNode) => void,
): void {
	for (const child of node.children) {
		if (REVISION_TAGS.has(child.tag)) visit(child);
		visitRevisions(child, visit);
	}
}

function revisionKey(revision: XmlNode): string {
	return [
		revision.tag,
		revision.getAttribute("w:id") ?? "",
		revision.getAttribute("w:author") ?? "",
		revision.getAttribute("w:date") ?? "",
	].join("\u0000");
}
