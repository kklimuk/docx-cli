import {
	isRunBearingWrapper,
	runTextLength,
	sliceRun,
	sumRunBearingTextLength,
	type XmlNode,
} from "../parser";

/** Maps a `--clear` attribute name to the `<w:rPr>` child tags it removes.
 *  Mutating rPr in place (rather than re-emitting via the AST) preserves any
 *  run property we don't model — the core in-place-mutation invariant. */
const CLEAR_TAG_MAP: Record<string, readonly string[]> = {
	bold: ["w:b", "w:bCs"],
	italic: ["w:i", "w:iCs"],
	strike: ["w:strike"],
	underline: ["w:u"],
	highlight: ["w:highlight"],
	shade: ["w:shd"],
	color: ["w:color"],
	font: ["w:rFonts"],
	size: ["w:sz", "w:szCs"],
	vertalign: ["w:vertAlign"],
	caps: ["w:caps"],
	smallcaps: ["w:smallCaps"],
	style: ["w:rStyle"],
};

/** Valid `--clear` attribute names (plus the special `all`). */
export const CLEARABLE_ATTRS: readonly string[] = Object.keys(CLEAR_TAG_MAP);

/** Resolve a list of `--clear` names (possibly including `all`) into the set of
 *  rPr child tags to strip. `all` expands to every formatting tag above —
 *  leaving non-formatting rPr children (e.g. `<w:lang>`) intact. Returns null
 *  if any name is unknown (caller reports the usage error). */
export function resolveClearTags(attrs: string[]): Set<string> | null {
	const names = attrs.includes("all") ? CLEARABLE_ATTRS : attrs;
	const tags = new Set<string>();
	for (const name of names) {
		const mapped = CLEAR_TAG_MAP[name];
		if (!mapped && name !== "all") return null;
		for (const tag of mapped ?? []) tags.add(tag);
	}
	return tags;
}

/** Strip the given formatting from the run(s) of `paragraph`. With `span`,
 *  only the runs (or slices of runs) overlapping `[start, end)` lose the
 *  formatting; without it, every run in the paragraph does. Offsets are
 *  accepted-view (matching `find`), so `find --highlight … | edit --clear …`
 *  lines up. Returns the number of runs whose rPr changed. */
export function clearFormatting(
	paragraph: XmlNode,
	span: { start: number; end: number } | null,
	tags: Set<string>,
): void {
	if (tags.size === 0) return;
	if (!span) {
		stripAllRuns(paragraph, tags);
		return;
	}
	clearInContainer(paragraph, span, tags, 0);
}

function stripAllRuns(node: XmlNode, tags: Set<string>): void {
	for (const child of node.children) {
		if (child.tag === "w:r") stripRunProperties(child, tags);
		else if (isRunBearingWrapper(child.tag)) stripAllRuns(child, tags);
	}
}

function clearInContainer(
	container: XmlNode,
	span: { start: number; end: number },
	tags: Set<string>,
	baseOffset: number,
): number {
	const out: XmlNode[] = [];
	let offset = baseOffset;
	for (const child of container.children) {
		if (child.tag === "w:r") {
			const length = runTextLength(child);
			const runStart = offset;
			const runEnd = offset + length;
			offset = runEnd;
			if (length === 0 || runEnd <= span.start || runStart >= span.end) {
				out.push(child);
				continue;
			}
			const sliceStart = Math.max(0, span.start - runStart);
			const sliceEnd = Math.min(length, span.end - runStart);
			if (sliceStart > 0) out.push(sliceRun(child, 0, sliceStart));
			const middle = sliceRun(child, sliceStart, sliceEnd);
			stripRunProperties(middle, tags);
			out.push(middle);
			if (sliceEnd < length) out.push(sliceRun(child, sliceEnd, length));
			continue;
		}
		// `<w:del>` / `<w:moveFrom>` are invisible in the accepted view — pass
		// through with no offset advance, matching `find`'s accounting.
		if (child.tag === "w:del" || child.tag === "w:moveFrom") {
			out.push(child);
			continue;
		}
		if (isRunBearingWrapper(child.tag)) {
			const innerLength = sumRunBearingTextLength(child.children);
			if (offset + innerLength <= span.start || offset >= span.end) {
				out.push(child);
				offset += innerLength;
				continue;
			}
			offset = clearInContainer(child, span, tags, offset);
			out.push(child);
			continue;
		}
		out.push(child);
	}
	container.children = out;
	return offset;
}

function stripRunProperties(run: XmlNode, tags: Set<string>): void {
	const rPr = run.findChild("w:rPr");
	if (!rPr) return;
	rPr.children = rPr.children.filter((child) => !tags.has(child.tag));
	if (rPr.children.length === 0) {
		run.children = run.children.filter((child) => child !== rPr);
	}
}
