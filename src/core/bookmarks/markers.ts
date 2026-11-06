import {
	isRunBearingWrapper,
	wrapperContentNode,
	type XmlNode,
} from "../parser";

/** Lift a paragraph's own `<w:bookmarkStart>`/`<w:bookmarkEnd>` markers out of
 *  it (directly or inside a run-bearing wrapper such as `<w:hyperlink>`),
 *  REMOVING them, so a whole-paragraph rebuild can't take them along —
 *  `edit --markdown` used to drop every bookmark on the paragraph, and with
 *  it every cross-reference that pointed there. Hand the result to
 *  {@link restoreBookmarkMarkers} once the new content is live. Markers inside
 *  a run (a text box's story) are not the paragraph's and stay with the run. */
export function liftBookmarkMarkers(paragraph: XmlNode): BookmarkMarkers {
	const markers = collectBookmarkMarkers(paragraph);
	const lifted = new Set([...markers.starts, ...markers.ends]);
	if (lifted.size > 0) removeMarkers(paragraph, lifted);
	return markers;
}

/** Snapshot the same markers WITHOUT removing them — for the in-place
 *  `--text` word diff, which keeps markers at their exact offsets unless the
 *  run container holding one (a `<w:hyperlink>`) is rebuilt away. */
export function collectBookmarkMarkers(paragraph: XmlNode): BookmarkMarkers {
	const markers: BookmarkMarkers = { starts: [], ends: [] };
	visitMarkerScope(paragraph, (child) => {
		if (child.tag === "w:bookmarkStart") markers.starts.push(child);
		if (child.tag === "w:bookmarkEnd") markers.ends.push(child);
	});
	return markers;
}

/** Put back every snapshotted marker that is no longer inside `paragraph`:
 *  a start right after the `<w:pPr>`, an end after the last child — so a
 *  bookmark that covered the paragraph covers its new content, and a
 *  cross-paragraph bookmark keeps the half that lived here (its partner in the
 *  neighbor stays paired). Markers still in place are left where they are. */
export function restoreBookmarkMarkers(
	paragraph: XmlNode,
	markers: BookmarkMarkers,
): void {
	// Keyed by tag + id, not node identity: a rebuild that CLONED a wrapper
	// (a split `<w:hyperlink>`) still carries the marker, just a copy of it.
	const current = collectBookmarkMarkers(paragraph);
	const present = new Set([...current.starts, ...current.ends].map(markerKey));
	const missing = (marker: XmlNode) => !present.has(markerKey(marker));
	const starts = markers.starts.filter(missing);
	const ends = markers.ends.filter(missing);
	if (starts.length === 0 && ends.length === 0) return;
	const insertAt = paragraph.children[0]?.tag === "w:pPr" ? 1 : 0;
	paragraph.children.splice(insertAt, 0, ...starts);
	paragraph.children.push(...ends);
}

function markerKey(marker: XmlNode): string {
	return `${marker.tag}#${marker.getAttribute("w:id") ?? ""}`;
}

/** Walk the paragraph's own marker scope — its children plus the contents of
 *  run-bearing wrappers (via `wrapperContentNode`, so an
 *  `<mc:AlternateContent>` is visited through its chosen branch only). */
function visitMarkerScope(
	node: XmlNode,
	visit: (child: XmlNode) => void,
): void {
	for (const child of node.children) {
		visit(child);
		if (isRunBearingWrapper(child.tag)) {
			visitMarkerScope(wrapperContentNode(child), visit);
		}
	}
}

/** Drop `markers` from the same scope {@link visitMarkerScope} walks. */
function removeMarkers(node: XmlNode, markers: Set<XmlNode>): void {
	// In place: a caller may hold this child array (a live block reference).
	for (let index = node.children.length - 1; index >= 0; index--) {
		const child = node.children[index];
		if (child && markers.has(child)) node.children.splice(index, 1);
	}
	for (const child of node.children) {
		if (isRunBearingWrapper(child.tag)) {
			removeMarkers(wrapperContentNode(child), markers);
		}
	}
}

/** The original marker NODES (not copies), so every attribute — `w:name`,
 *  `w:colFirst`/`w:colLast`, `w:displacedByCustomXml` — survives the move. */
export type BookmarkMarkers = { starts: XmlNode[]; ends: XmlNode[] };
