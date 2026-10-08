import { isRevisionKindVisible, type RevisionView } from "../parser";
import type { TrackedChange } from "./types";

export type { RevisionView };

/**
 * Whether content inside `change` shows in `view`. It does only when EVERY
 * enclosing revision does — the innermost wrapper and each one it nests in
 * (`TrackedChange.within`): rejecting an insertion removes everything inside
 * it, including another author's deletion nested there, and accepting a
 * deletion removes an insertion nested inside it. The per-kind rule is
 * `isRevisionKindVisible` (parser/revision-view.ts), which the XML offset
 * walkers share via `isWrapperVisibleInView` — they recurse only into visible
 * wrappers — so read, find, wc and replace offsets stay in sync. Content
 * outside any revision always shows.
 */
export function isRevisionVisible(
	change: TrackedChange | undefined,
	view: RevisionView,
): boolean {
	for (let revision = change; revision; revision = revision.within) {
		if (!isRevisionKindVisible(revision.kind, view)) return false;
	}
	return true;
}
