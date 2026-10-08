import {
	isRunBearingWrapper,
	isTrackedChangeWrapper,
	runTextLength,
	wrapperContent,
} from "./run-ops";
import type { XmlNode } from "./xml-node";

/**
 * THE view-visibility rule for tracked changes — the one place it is stated.
 * Content inside a revision of `kind` shows in `view` unless that view
 * resolves the revision away: accepting drops deletions and move sources,
 * rejecting drops insertions and move destinations; `current` shows all.
 * Both sides of the AST↔XML offset bridge derive from it: the AST walks a
 * run's `TrackedChange.within` chain through `isRevisionVisible`
 * (core/ast/revision-visibility.ts), and the XML walkers recurse only into
 * wrappers `isWrapperVisibleInView` passes — so nested content counts only
 * when EVERY enclosing revision is visible, on both sides alike.
 */
export function isRevisionKindVisible(
	kind: string,
	view: RevisionView,
): boolean {
	return view === "current" || !HIDDEN_REVISION_KINDS[view].has(kind);
}

const HIDDEN_REVISION_KINDS: Record<
	Exclude<RevisionView, "current">,
	ReadonlySet<string>
> = {
	accepted: new Set(["del", "moveFrom"]),
	baseline: new Set(["ins", "moveTo"]),
};

/** Whether a run-bearing wrapper's contents are visible in `view`. A
 *  revision wrapper (`<w:ins>`/`<w:del>`/`<w:moveFrom>`/`<w:moveTo>`) follows
 *  its kind; every other run-bearing wrapper (a link, a field, a smart tag,
 *  an AlternateContent) is transparent. Not a run-bearing wrapper → false. */
export function isWrapperVisibleInView(
	tag: string,
	view: RevisionView,
): boolean {
	if (!isRunBearingWrapper(tag)) return false;
	if (!isTrackedChangeWrapper(tag)) return true;
	return isRevisionKindVisible(tag.slice("w:".length), view);
}

/** Text length of `children` in `view`: every run, descending only into
 *  wrappers visible there — the XML-side twin of the AST's view text, so a
 *  walker that SKIPS a wrapper skips exactly the offset `find` counted. */
export function sumVisibleTextLength(
	children: XmlNode[],
	view: RevisionView,
): number {
	let total = 0;
	for (const child of children) {
		if (child.tag === "w:r") {
			total += runTextLength(child);
			continue;
		}
		if (isWrapperVisibleInView(child.tag, view)) {
			total += sumVisibleTextLength(wrapperContent(child), view);
		}
	}
	return total;
}

/** Which resolution of the document's tracked changes is being shown:
 *  `current` (everything, marked up), `accepted` (as if all were accepted),
 *  `baseline` (as if all were rejected). */
export type RevisionView = "current" | "accepted" | "baseline";
