import type { RunFormat } from "../../edit/set-formatting";
import {
	isSubtractiveTrackedChangeWrapper,
	wrapperContentNode,
	type XmlNode,
} from "../../parser";
import type { RevisionAllocator, TrackedMeta } from "../../track-changes";
import type { FindView } from "../index";
import { chooseLevel } from "./level";
import { rebuildAroundSpan } from "./rebuild";
import { anchorFor, type PathEntry, spanTarget } from "./target";

export { replacementRuns } from "./replacement-runs";

/**
 * Replace text in a paragraph's runs at the given span with `replacement`.
 * Surrounding text and run formatting are preserved; the replacement run
 * inherits the rPr of the first run that overlaps the span (for an empty span
 * at a run boundary — a pure insertion, `pN:S-S` — the run it is anchored on).
 *
 * The span uses paragraph-relative offsets in `view`, matching the AST's
 * accounting (runs nested in visible wrappers count; hidden revisions don't).
 *
 * The edit is rebuilt at ONE level — a container's child list — and every
 * wrapper there that the span crosses is split around it, recursively, so
 * attribution survives on both sides (`rebuildAroundSpan`). The level is the
 * deepest container holding the whole span (a hyperlink the span sits in, so
 * the replacement inherits the link), except:
 *
 * - Under tracking (or for an empty span), never inside a revision wrapper
 *   (`<w:ins>`/`<w:del>`/`<w:moveFrom>`/`<w:moveTo>`) at ANY depth (issue
 *   #13): the level rises to the container above the OUTERMOST revision on
 *   the span's path, and every wrapper below splits — so the replacement's own
 *   `<w:ins>` is never a descendant of another revision (inside one it would
 *   be attributed to that author, or vanish when that revision resolves). The
 *   cut's `<w:del>` stays nested where the text was (Word's shape); cut text
 *   an enclosing deletion already removed stays as it is. A link the span sat
 *   in is re-created around the replacement, so it keeps the link.
 * - The tracked author's OWN pending insertion (the nearest enclosing
 *   revision, an `<w:ins>`/`<w:moveTo>` whose `w:author` matches exactly, with
 *   no other revision between it and the span's runs) is not a new revision:
 *   the span is rebuilt untracked inside it — the cut vanishes and the
 *   replacement joins that insertion (same `w:id`/date), as Word does when you
 *   edit your own pending text.
 */
export function replaceSpanInParagraph(
	paragraph: XmlNode,
	span: Span,
	replacement: string,
	tracked?: TrackedReplaceOptions,
	view: FindView = "accepted",
	formatting?: ReplacementFormatting,
): void {
	if (span.start > span.end) {
		throw new Error(
			`replaceSpanInParagraph: invalid span ${span.start}-${span.end}`,
		);
	}

	const target = spanTarget(paragraph, span, view);
	const level = chooseLevel(target, span, tracked);
	const depth = level.path.length;
	const container = contentAt(paragraph, level.path, depth);
	const anchor = anchorFor(target, depth, level.linkShell);

	rebuildAroundSpan(container, level.path.at(-1)?.start ?? 0, {
		span,
		view,
		tracked: level.ownInsertion ? null : (tracked ?? null),
		deleted: level.path.some((entry) =>
			isSubtractiveTrackedChangeWrapper(entry.wrapper.tag),
		),
		replacement: {
			text: replacement,
			runProperties: anchor?.run.findChild("w:rPr")?.clone() ?? null,
			formatting,
			linkShell: level.linkShell,
		},
	});

	if (level.ownInsertion) {
		const index = level.path.findIndex(
			(entry) => entry.wrapper === level.ownInsertion,
		);
		dropIfBlank(level.ownInsertion, contentAt(paragraph, level.path, index));
	}
}

/** The node whose child list holds the content at `depth` along `path`: the
 *  paragraph itself at depth 0, else the `depth`-th wrapper's content. */
function contentAt(
	paragraph: XmlNode,
	path: PathEntry[],
	depth: number,
): XmlNode {
	const entry = path[depth - 1];
	return entry ? wrapperContentNode(entry.wrapper) : paragraph;
}

/** Remove an own insertion the merge left holding nothing but blank runs —
 *  deleting all of your own pending text leaves no revision behind, as in
 *  Word, rather than a phantom empty `<w:ins>` that `track-changes list`
 *  would still report. */
function dropIfBlank(wrapper: XmlNode, container: XmlNode): void {
	if (!wrapper.children.every(isBlankRun)) return;
	const index = container.children.indexOf(wrapper);
	if (index >= 0) container.children.splice(index, 1);
}

function isBlankRun(node: XmlNode): boolean {
	return (
		node.tag === "w:r" &&
		node.children.every(
			(child) =>
				child.tag === "w:rPr" ||
				(child.tag === "w:t" && child.collectText().length === 0),
		)
	);
}

export type Span = { start: number; end: number };

export type TrackedReplaceOptions = {
	meta: Omit<TrackedMeta, "revisionId">;
	allocator: RevisionAllocator;
};

export type ReplacementFormatting = {
	clearTags?: Set<string>;
	format?: RunFormat;
};
