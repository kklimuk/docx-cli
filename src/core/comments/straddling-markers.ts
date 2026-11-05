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
	const kindsById = new Map<string, Set<string>>();
	for (const marker of markers) {
		const kinds = kindsById.get(marker.id) ?? new Set<string>();
		kinds.add(marker.kind);
		kindsById.set(marker.id, kinds);
	}
	return markers
		.filter((marker) => (kindsById.get(marker.id)?.size ?? 0) < 3)
		.map((marker) => marker.node);
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
