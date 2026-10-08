import type { TrackedChange, TrackedChangeKind } from "./types";

/**
 * Whether content inside `change` shows in `view`. It does only when EVERY
 * enclosing revision does — the innermost wrapper and each one it nests in
 * (`TrackedChange.within`): rejecting an insertion removes everything inside
 * it, including another author's deletion nested there, and accepting a
 * deletion removes an insertion nested inside it. This is the AST-side twin
 * of the XML offset walkers, which only recurse into visible wrappers
 * (`isWrapperVisibleInView` in find/replace-span.tsx), so read, find, wc and
 * replace offsets stay in sync. Content outside any revision always shows.
 */
export function isRevisionVisible(
	change: TrackedChange | undefined,
	view: RevisionView,
): boolean {
	if (view === "current") return true;
	const hidden = HIDDEN_KINDS[view];
	for (let revision = change; revision; revision = revision.within) {
		if (hidden.has(revision.kind)) return false;
	}
	return true;
}

const HIDDEN_KINDS: Record<
	"accepted" | "baseline",
	ReadonlySet<TrackedChangeKind>
> = {
	accepted: new Set(["del", "moveFrom"]),
	baseline: new Set(["ins", "moveTo"]),
};

/** Which resolution of the document's tracked changes is being shown:
 *  `current` (everything, marked up), `accepted` (as if all were accepted),
 *  `baseline` (as if all were rejected). */
export type RevisionView = "current" | "accepted" | "baseline";
