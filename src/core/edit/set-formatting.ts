import { applyRunFont } from "../ast/document/styles";
import { insertRprChildInOrder } from "../blocks";
import {
	isRunBearingWrapper,
	runTextLength,
	sliceRun,
	sumRunBearingTextLength,
	XmlNode,
} from "../parser";

/** The run properties `setFormatting` writes — the structural inverse of the
 *  `--clear` vocabulary (clear REMOVES these `<w:rPr>` children, set ADDS or
 *  replaces them). Each present field is applied; an absent field is left
 *  untouched. The boolean toggles only turn a property ON — to turn one OFF use
 *  `--clear` (which strips the `<w:rPr>` child). We set rPr in place rather than
 *  re-emitting the run, so any run property we don't model survives — the same
 *  in-place-mutation invariant `clearFormatting` relies on. */
export type RunFormat = {
	bold?: boolean;
	italic?: boolean;
	strike?: boolean;
	allCaps?: boolean;
	smallCaps?: boolean;
	/** `ST_Underline` style (e.g. `single`, `double`, `wave`). */
	underline?: string;
	/** Hex color riding on the `<w:u>` element (no `#`). */
	underlineColor?: string;
	/** Run text color — 6-digit hex, no `#`. */
	color?: string;
	/** `ST_HighlightColor` name (the 16-color highlighter palette). */
	highlight?: string;
	/** Arbitrary hex background fill (no `#`) — distinct from the highlight palette. */
	shade?: string;
	font?: string;
	/** Font size in half-points (24 = 12pt). */
	sizeHalfPoints?: number;
	/** `superscript` | `subscript`. */
	vertAlign?: string;
};

/** Apply `format` to the run(s) of `paragraph`. With `span`, only the runs (or
 *  slices of runs) overlapping `[start, end)` are affected; without it, every
 *  run in the paragraph is. Offsets are accepted-view (matching `find`), so
 *  `find … | edit --bold` lines up. The exact structural mirror of
 *  `clearFormatting`: it walks runs the same way — slicing a partially covered
 *  run into pre / middle / post at the span boundaries — but ADDS/replaces rPr
 *  children on the middle slice instead of stripping them. */
export function setFormatting(
	paragraph: XmlNode,
	span: { start: number; end: number } | null,
	format: RunFormat,
): void {
	if (!hasAnyFormat(format)) return;
	if (!span) {
		setAllRuns(paragraph, format);
		return;
	}
	setInContainer(paragraph, span, format, 0);
}

function hasAnyFormat(format: RunFormat): boolean {
	return Object.values(format).some((value) => value !== undefined);
}

function setAllRuns(node: XmlNode, format: RunFormat): void {
	for (const child of node.children) {
		if (child.tag === "w:r") setRunProperties(child, format);
		else if (isRunBearingWrapper(child.tag)) setAllRuns(child, format);
	}
}

function setInContainer(
	container: XmlNode,
	span: { start: number; end: number },
	format: RunFormat,
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
			setRunProperties(middle, format);
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
			offset = setInContainer(child, span, format, offset);
			out.push(child);
			continue;
		}
		out.push(child);
	}
	container.children = out;
	return offset;
}

/** Add/replace each named property as a `<w:rPr>` child, find-or-creating the
 *  `<w:rPr>` first. Children are spliced at their canonical CT_RPr slot (never
 *  pushed) so Word accepts the run. Replacing (drop any existing same-tag child
 *  first) means a SET overrides a prior value rather than duplicating it. Font
 *  reuses `applyRunFont` so an explicit family beats a theme reference (and
 *  East-Asian fallback survives); size sets both `<w:sz>` and `<w:szCs>` so
 *  complex-script text resizes too. */
function setRunProperties(run: XmlNode, format: RunFormat): void {
	let rPr = run.findChild("w:rPr");
	if (!rPr) {
		rPr = XmlNode.element("w:rPr");
		// rPr must be the FIRST child of a `<w:r>` (it precedes `<w:t>`).
		run.children.unshift(rPr);
	}

	if (format.font !== undefined) {
		let rFonts = rPr.findChild("w:rFonts");
		if (!rFonts) {
			rFonts = XmlNode.element("w:rFonts");
			insertRprChildInOrder(rPr, rFonts);
		}
		applyRunFont(rFonts, format.font);
	}
	if (format.bold) putToggle(rPr, "w:b");
	if (format.italic) putToggle(rPr, "w:i");
	if (format.allCaps) putToggle(rPr, "w:caps");
	if (format.smallCaps) putToggle(rPr, "w:smallCaps");
	if (format.strike) putToggle(rPr, "w:strike");
	if (format.color !== undefined) putValue(rPr, "w:color", format.color);
	if (format.sizeHalfPoints !== undefined) {
		const value = String(format.sizeHalfPoints);
		putValue(rPr, "w:sz", value);
		putValue(rPr, "w:szCs", value);
	}
	if (format.highlight !== undefined)
		putValue(rPr, "w:highlight", format.highlight);
	if (format.underline !== undefined) {
		const u = XmlNode.element("w:u");
		u.setAttribute("w:val", format.underline);
		if (format.underlineColor !== undefined)
			u.setAttribute("w:color", format.underlineColor);
		putChild(rPr, u);
	}
	if (format.shade !== undefined) {
		const shd = XmlNode.element("w:shd");
		shd.setAttribute("w:val", "clear");
		shd.setAttribute("w:color", "auto");
		shd.setAttribute("w:fill", format.shade);
		putChild(rPr, shd);
	}
	if (format.vertAlign !== undefined)
		putValue(rPr, "w:vertAlign", format.vertAlign);
}

/** Drop any existing child with `child`'s tag, then splice `child` in at its
 *  canonical CT_RPr position. */
function putChild(rPr: XmlNode, child: XmlNode): void {
	rPr.children = rPr.children.filter((existing) => existing.tag !== child.tag);
	insertRprChildInOrder(rPr, child);
}

/** Replace-or-insert an empty toggle element (`<w:b/>`, `<w:i/>`, …). */
function putToggle(rPr: XmlNode, tag: string): void {
	putChild(rPr, XmlNode.element(tag));
}

/** Replace-or-insert a single-attribute element (`<w:color w:val=…/>`, etc.). */
function putValue(rPr: XmlNode, tag: string, value: string): void {
	const element = XmlNode.element(tag);
	element.setAttribute("w:val", value);
	putChild(rPr, element);
}
