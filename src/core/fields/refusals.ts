import type { Body } from "../ast/document/body";
import type { XmlNode } from "../parser";
import {
	comparePositions,
	type Field,
	type FieldPosition,
	fieldParagraphs,
	fieldType,
	holdsFieldMarkup,
	isFieldVisible,
	type ParsedFields,
	parseFields,
} from "./parse";

/** A field is atomic structure: a mutation carries it through intact, removes
 *  it whole, or refuses before writing anything (issue #42). These are the
 *  three refusals, each `null` when the mutation is safe. */

/** Refuse a whole-paragraph rebuild (`edit` markdown/`--text`/`--runs`/code,
 *  a `pN-pM` range edit) of paragraphs that hold a field: the rebuild emits
 *  fresh runs from text, so every field in them would come back as frozen
 *  plain text — instruction gone, exit 0, `read` unchanged. */
export function paragraphFieldRefusal(
	parsed: ParsedFields,
	blocks: XmlNode[],
	locator: string,
): FieldRefusal | null {
	const targets = paragraphsIn(blocks);
	const held = parsed.fields.filter(
		(field) =>
			isFieldVisible(field) &&
			[...fieldParagraphs(field)].some((paragraph) => targets.has(paragraph)),
	);
	if (held.length === 0) return null;
	const count = `${held.length} field${held.length === 1 ? "" : "s"}`;
	return {
		code: "FIELD_IN_PARAGRAPH",
		message: `${locator} contains ${count} (${describeFields(held)}) that this edit would turn into plain text. Nothing was changed.`,
		hint: `A field (cross-reference, page number, TOC entry, merge field, …) is kept only by edits around it: change the text beside it with a span edit (\`docx find FILE "TEXT"\` gives the pN:S-E locator for \`docx edit FILE --at pN:S-E "…"\`) or with \`docx replace\` — fields outside the span are kept, a span inside a field's result edits that result, and a span covering a whole field deletes it.`,
	};
}

/** Refuse a span that CUTS a field: one covering some of {begin … end} but
 *  not all of it. A span strictly inside the result (edits the cached text)
 *  and one covering the whole field (removes it — see `removeCoveredFields`)
 *  are both fine. The `equationSpanRefusal` analog. */
export function fieldSpanRefusal(
	parsed: ParsedFields,
	start: SpanEnd,
	end: SpanEnd,
	locator: string,
): FieldRefusal | null {
	const span = positionSpan(parsed, start, end);
	if (!span) return null;
	const cut = parsed.fields.find(
		(field) => isFieldVisible(field) && spanRelation(field, span) === "cut",
	);
	if (!cut) return null;
	return {
		code: "UNSUPPORTED",
		message: `${locator}: the span cuts a field (${describeFields([cut])}) — it covers part of the field but not all of it, which would leave half a field behind. Nothing was changed.`,
		hint: `A field is one unit whose visible text is its result${cut.result ? ` (${JSON.stringify(cut.result)})` : ""}. Match only text beside it, only text inside its result to edit that result, or the whole result plus text on either side to delete the field.`,
	};
}

/** Refuse removing paragraphs (`delete --at pN`/`pN-pM`, `edit --at pN
 *  --text ""`, a batch's line removals) that hold one end of a field but not
 *  the other — a TOC-shaped field that opens in one paragraph and closes in
 *  another. Removing a paragraph that holds a whole field is fine: the field
 *  goes with it. Positions don't matter here, so any view's parse will do. */
export function fieldRemovalRefusal(
	body: Body,
	removedBlocks: XmlNode[],
): FieldRefusal | null {
	if (!holdsFieldMarkup(removedBlocks)) return null;
	const removed = paragraphsIn(removedBlocks);
	const locatorOf = (paragraph: XmlNode): string =>
		body.findBlockByNode(paragraph)?.id ?? "a paragraph";
	for (const field of parseFields([body.body]).fields) {
		const paragraphs = [...fieldParagraphs(field)];
		const taken = paragraphs.filter((paragraph) => removed.has(paragraph));
		if (taken.length === 0 || taken.length === paragraphs.length) continue;
		const first = paragraphs[0];
		const last = paragraphs.at(-1);
		const extent =
			first && last ? `${locatorOf(first)}–${locatorOf(last)}` : "";
		return {
			code: "UNSUPPORTED",
			message: `Removing ${taken.map(locatorOf).join(", ")} would split a field (${describeFields([field])}) that spans ${extent}: the removal takes one end of it and leaves the other, half a field. Nothing was changed.`,
			hint: `Remove every paragraph the field spans in one call (\`docx delete FILE --at ${extent.replace("–", "-")}\`), or change the text inside it with a span edit instead.`,
		};
	}
	return null;
}

/** A span given as paragraph-local offsets (`pN:S-E`, or across paragraphs
 *  `pN:S-pM:E`), in the positions `parsed` measured — null when a paragraph
 *  wasn't walked (it's in no story `parsed` covers). */
export function positionSpan(
	parsed: ParsedFields,
	start: SpanEnd,
	end: SpanEnd,
): PositionSpan | null {
	const startOrdinal = parsed.ordinals.get(start.paragraph);
	const endOrdinal = parsed.ordinals.get(end.paragraph);
	if (startOrdinal === undefined || endOrdinal === undefined) return null;
	return {
		start: { ordinal: startOrdinal, offset: start.offset },
		end: { ordinal: endOrdinal, offset: end.offset },
	};
}

/** How a span meets a field, positions compared in document order. A field
 *  occupies [begin, end] (zero-width markers) and its result [separate, end].
 *  `inside` wins a tie: a span exactly over the result edits it, keeping the
 *  field — the walkers leave a marker sitting on a span edge in place. */
export function spanRelation(
	field: Field,
	span: PositionSpan,
): "disjoint" | "inside" | "covered" | "cut" {
	const begin = field.kind === "simple" ? field.start : field.begin.position;
	const end = field.kind === "simple" ? field.end : field.end.position;
	const resultStart =
		field.kind === "simple"
			? field.start
			: (field.separate?.position ?? field.begin.position);
	const overlaps =
		comparePositions(span.start, end) < 0 &&
		comparePositions(span.end, begin) > 0;
	if (!overlaps) return "disjoint";
	if (
		comparePositions(resultStart, span.start) <= 0 &&
		comparePositions(span.end, end) <= 0
	) {
		return "inside";
	}
	if (
		comparePositions(span.start, begin) <= 0 &&
		comparePositions(end, span.end) <= 0
	) {
		return "covered";
	}
	return "cut";
}

/** `REF ×3: "1", "3", "4]"` — each kind with a count and the cached results,
 *  the text an agent sees in `read`. */
function describeFields(fields: Field[]): string {
	const byType = new Map<string, Field[]>();
	for (const field of fields) {
		const type = fieldType(field);
		byType.set(type, [...(byType.get(type) ?? []), field]);
	}
	return [...byType]
		.map(([type, group]) => {
			const count = group.length > 1 ? ` ×${group.length}` : "";
			const results = group
				.map((field) => JSON.stringify(truncate(field.result)))
				.join(", ");
			return `${type}${count}: ${results}`;
		})
		.join("; ");
}

function truncate(text: string): string {
	return text.length > 40 ? `${text.slice(0, 39)}…` : text;
}

/** Every `<w:p>` in (or being) one of `blocks` — a removed or rebuilt table
 *  takes its cells' paragraphs with it. A text box's story is NOT the anchor
 *  paragraph's content: removal re-anchors the box and a rebuild lifts it out
 *  and puts it back, so its paragraphs (and their fields) survive both. */
function paragraphsIn(blocks: XmlNode[]): Set<XmlNode> {
	const out = new Set<XmlNode>();
	const visit = (node: XmlNode): void => {
		if (node.tag === "w:txbxContent") return;
		if (node.tag === "w:p") out.add(node);
		for (const child of node.children) {
			if (!child.isText) visit(child);
		}
	};
	for (const block of blocks) visit(block);
	return out;
}

export type PositionSpan = { start: FieldPosition; end: FieldPosition };

/** One end of a span: a paragraph and a paragraph-local offset. */
export type SpanEnd = { paragraph: XmlNode; offset: number };

export type FieldRefusal = {
	code: "FIELD_IN_PARAGRAPH" | "UNSUPPORTED";
	message: string;
	hint: string;
};
