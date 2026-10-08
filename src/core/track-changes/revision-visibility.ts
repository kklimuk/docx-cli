import type { TrackedChange, TrackedChangeKind } from "../ast/types";

/** Whether a run carrying `change` is visible in `view` — the ONE rule every
 *  view consumer shares (`paragraphTextAccepted`/`paragraphTextBaseline`, so
 *  `wc` follows; `find`'s per-view text; `iterateBlocks({ view })`; the
 *  markdown renderer). `accepted` hides `del`/`moveFrom`, `baseline` hides
 *  `ins`/`moveTo`, `current` shows everything — and a run is visible only when
 *  EVERY enclosing wrapper is: a `<w:del>` nested inside another author's
 *  `<w:ins>` is hidden in BOTH resolved views, because rejecting the insertion
 *  removes everything inside it (the deletion included) and accepting it still
 *  leaves the deletion to be accepted. The XML-side length/placement walkers
 *  (`sumVisibleTextLength` in find/replace-span.tsx, find/replace-across.tsx
 *  and comments/markers.tsx) never descend a hidden wrapper, so they follow
 *  the same rule structurally. */
export function isRevisionVisible(
	change: TrackedChange | undefined,
	view: RevisionView,
): boolean {
	if (view === "current" || !change) return true;
	if (!isKindVisible(change.kind, view)) return false;
	for (const outer of change.within ?? []) {
		if (!isKindVisible(outer.kind, view)) return false;
	}
	return true;
}

export type RevisionView = "current" | "accepted" | "baseline";

function isKindVisible(
	kind: TrackedChangeKind,
	view: Exclude<RevisionView, "current">,
): boolean {
	if (view === "accepted") return kind !== "del" && kind !== "moveFrom";
	return kind !== "ins" && kind !== "moveTo";
}
