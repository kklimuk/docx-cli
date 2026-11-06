import { isRunBearingWrapper, wrapperContent, XmlNode } from "../parser";

/** The comment markers in `children` (range starts/ends and reference runs,
 *  searched through run-bearing wrappers) whose comment ALSO has a marker
 *  outside them — in document order, ready to splice back in when `children`
 *  is about to be removed. A comment wholly inside (start, end AND reference)
 *  is left out: it goes with its text. Dropping a straddling marker with the
 *  content strands the other half — a start with no end, a range with no
 *  reference — which `validate` reports and Word drops. Reference runs come
 *  back as fresh runs carrying only their `<w:rPr>` + `<w:commentReference>`. */
export function straddlingCommentMarkers(children: XmlNode[]): XmlNode[] {
	const markers = commentMarkersWithin(children);
	const kindsById = markerKindsById(markers);
	return markers
		.filter((marker) => (kindsById.get(marker.id)?.size ?? 0) < 3)
		.map((marker) => marker.node);
}

/** The ids of the comments WHOLLY inside `children` — range start, range end
 *  AND reference all within, searched through run-bearing wrappers — the
 *  complement of `straddlingCommentMarkers`. Removing `children` removes the
 *  whole anchor, so the comment itself must go too: a `<w:comment>` with no
 *  marker anywhere is what `validate` reports as "no anchor" and Word
 *  silently deletes on open. */
export function containedCommentIds(children: XmlNode[]): string[] {
	return [...markerKindsById(commentMarkersWithin(children))]
		.filter(([, kinds]) => kinds.size === 3)
		.map(([id]) => id);
}

/** Which marker kinds (start, end, reference) each comment id has among
 *  `markers` — 3 means the comment lies wholly inside them. */
function markerKindsById(markers: CommentMarker[]): Map<string, Set<string>> {
	const kindsById = new Map<string, Set<string>>();
	for (const marker of markers) {
		const kinds = kindsById.get(marker.id) ?? new Set<string>();
		kinds.add(marker.kind);
		kindsById.set(marker.id, kinds);
	}
	return kindsById;
}

function commentMarkersWithin(children: XmlNode[]): CommentMarker[] {
	const markers: CommentMarker[] = [];
	for (const child of children) {
		if (
			child.tag === "w:commentRangeStart" ||
			child.tag === "w:commentRangeEnd"
		) {
			const id = child.getAttribute("w:id");
			if (id !== undefined) markers.push({ id, kind: child.tag, node: child });
			continue;
		}
		if (child.tag === "w:r") {
			const reference = child.findChild("w:commentReference");
			const id = reference?.getAttribute("w:id");
			if (!reference || id === undefined) continue;
			const run = new XmlNode("w:r", { ...child.attributes });
			const rPr = child.findChild("w:rPr");
			run.children = rPr ? [rPr, reference] : [reference];
			markers.push({ id, kind: reference.tag, node: run });
			continue;
		}
		if (isRunBearingWrapper(child.tag)) {
			markers.push(...commentMarkersWithin(wrapperContent(child)));
		}
	}
	return markers;
}

type CommentMarker = { id: string; kind: string; node: XmlNode };
