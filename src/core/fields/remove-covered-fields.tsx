import type { FindView } from "../find";
import { XmlNode } from "../parser";
import { convertTextToDelText, type TrackedMeta } from "../track-changes";
import { Del } from "../track-changes/emit";
import {
	type FieldMarker,
	fieldMarkers,
	type ParsedFields,
	parseFields,
} from "./parse";
import { positionSpan, type SpanEnd, spanRelation } from "./refusals";

/** `removeCoveredFields` for a span inside one paragraph (`pN:S-E`, a
 *  `replace` match): a field the span covers whole has both ends in this
 *  paragraph, so the paragraph alone is parsed — LIFO pairing makes its local
 *  pairs exactly the document's. */
export function removeFieldsCoveredInParagraph(
	paragraph: XmlNode,
	span: { start: number; end: number },
	view: FindView,
	mintMeta?: () => TrackedMeta,
): void {
	const story = new XmlNode("w:body");
	story.children = [paragraph];
	removeCoveredFields(
		parseFields([story], view),
		{ paragraph, offset: span.start },
		{ paragraph, offset: span.end },
		mintMeta,
	);
}

/** Before a span walker deletes `span`, take out the markers of every complex
 *  field the span COVERS whole, so the field goes as one unit. The walkers
 *  own text, not structure: a zero-width `<w:fldChar>` sitting on a span edge
 *  stays where it is, so a span from before a field to the end of its result
 *  took the begin and separate and left the `end` behind — an orphan
 *  `validate` couldn't see (issue #42). The result runs are text inside the
 *  span; the walker deletes those itself.
 *
 *  Untracked, the markers are dropped. Tracked (`mintMeta`), each marker's run
 *  is wrapped in its own `<w:del>` (`convertTextToDelText` renames
 *  `<w:instrText>` to `<w:delInstrText>`; reject renames it back), so Accept All removes the
 *  field whole and Reject All restores it. A marker sharing its run with other
 *  content is first moved into a run of its own. Markers already hidden in a
 *  pending deletion are left alone. A covered `<w:fldSimple>` needs nothing:
 *  the walker removes it as the wrapper of the text it deletes. */
export function removeCoveredFields(
	parsed: ParsedFields,
	start: SpanEnd,
	end: SpanEnd,
	mintMeta?: () => TrackedMeta,
): void {
	const span = positionSpan(parsed, start, end);
	if (!span) return;
	const markers = parsed.fields.flatMap((field) =>
		field.kind === "complex" && spanRelation(field, span) === "covered"
			? fieldMarkers(field).filter((marker) => marker.visible)
			: [],
	);
	const byRun = new Map<XmlNode, FieldMarker[]>();
	for (const marker of markers) {
		byRun.set(marker.run, [...(byRun.get(marker.run) ?? []), marker]);
	}
	for (const [run, runMarkers] of byRun) {
		const [first] = runMarkers;
		if (!first) continue;
		const target = isolateMarkers(
			run,
			first.container,
			new Set(runMarkers.map((marker) => marker.node)),
		);
		const index = first.container.children.indexOf(target);
		if (index === -1) continue;
		if (!mintMeta) {
			first.container.children.splice(index, 1);
			continue;
		}
		first.container.children.splice(
			index,
			1,
			<Del meta={mintMeta()}>{convertTextToDelText(target)}</Del>,
		);
	}
}

/** The run holding exactly `nodes` (plus its rPr): `run` itself when they're
 *  all it holds, else a new run — same attributes and rPr — that they're moved
 *  into, placed just before `run` in `container`. */
function isolateMarkers(
	run: XmlNode,
	container: XmlNode,
	nodes: Set<XmlNode>,
): XmlNode {
	const content = run.children.filter(
		(child) => !child.isText && child.tag !== "w:rPr",
	);
	if (content.every((child) => nodes.has(child))) return run;
	const isolated = new XmlNode("w:r", { ...run.attributes });
	const properties = run.findChild("w:rPr");
	if (properties) isolated.children.push(properties.clone());
	isolated.children.push(...content.filter((child) => nodes.has(child)));
	run.children = run.children.filter((child) => !nodes.has(child));
	const index = container.children.indexOf(run);
	container.children.splice(Math.max(index, 0), 0, isolated);
	return isolated;
}
