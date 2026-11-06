import { XmlNode } from "../parser";
import { computeExtentEmu, type ImageSource } from "./source";

/** Size the frame of one existing drawing for a REPLACEMENT image, so the new
 *  picture is never stretched into the old picture's box (issue #46). Pure —
 *  reads the drawing's current `<wp:extent>`, mutates nothing.
 *
 *  - `keepSize`: the old box, verbatim (the caller asserts same shape).
 *  - `--width`/`--height`: `computeExtentEmu`, exactly as `images add` sizes
 *    (one axis keeps the new aspect ratio, both are honored as given).
 *  - default: keep the DISPLAYED width, recompute the height from the new
 *    image's pixel aspect ratio, then scale down (aspect kept) to fit the text
 *    area — so a replace can't introduce `overflow="yes"`, and a tall picture
 *    swapped into a wide banner can't grow taller than the page.
 *  - new image has no parsed pixel size (emf/wmf/webp/…): keep the old box and
 *    report `sizeKnown: false` so the caller can say so instead of silently
 *    distorting.
 *
 *  Returns null when a single-axis override needs an aspect ratio the new image
 *  doesn't expose — the caller asks for both `--width` and `--height`. */
export function fitReplacementExtent(
	drawing: XmlNode,
	source: Pick<ImageSource, "pixelWidth" | "pixelHeight">,
	sizing: ReplacementSizing,
	textArea: { widthEmu: number; heightEmu: number },
): ReplacementExtent | null {
	const current = currentExtent(drawing);
	if (sizing.keepSize) return { ...current, sizeKnown: true };

	if (sizing.widthInches !== undefined || sizing.heightInches !== undefined) {
		const extent = computeExtentEmu(source, sizing, textArea.widthEmu);
		return extent ? { ...extent, sizeKnown: true } : null;
	}

	const { pixelWidth, pixelHeight } = source;
	if (!pixelWidth || !pixelHeight) return { ...current, sizeKnown: false };
	// A frame with no readable width falls back to the native size, like `add`.
	const extent =
		current.widthEmu > 0
			? {
					widthEmu: current.widthEmu,
					heightEmu: Math.round((current.widthEmu * pixelHeight) / pixelWidth),
				}
			: computeExtentEmu(source, {});
	if (!extent) return null;
	return { ...fitWithin(extent, textArea), sizeKnown: true };
}

/** Scale an extent down (never up), aspect kept, until it fits `bounds`. A
 *  non-positive bound is treated as unknown and doesn't constrain. */
function fitWithin(
	extent: { widthEmu: number; heightEmu: number },
	bounds: { widthEmu: number; heightEmu: number },
): { widthEmu: number; heightEmu: number } {
	const ratio = (limit: number, size: number) =>
		limit > 0 && size > limit ? limit / size : 1;
	const scale = Math.min(
		ratio(bounds.widthEmu, extent.widthEmu),
		ratio(bounds.heightEmu, extent.heightEmu),
	);
	if (scale === 1) return extent;
	return {
		widthEmu: Math.round(extent.widthEmu * scale),
		heightEmu: Math.round(extent.heightEmu * scale),
	};
}

/** Apply a replacement frame to a drawing IN PLACE: both extents (`<wp:extent>`
 *  on the inline/anchor and `<a:ext>` in `pic:spPr/a:xfrm`, which must agree),
 *  drop any `<a:srcRect>` crop (it's in the OLD image's coordinates and would
 *  mis-crop the new one), and reset a floating image's tight/through
 *  `<wp:wrapPolygon>` to the full frame (it traced the OLD picture's outline). */
export function applyReplacementExtent(
	drawing: XmlNode,
	extent: { widthEmu: number; heightEmu: number },
): void {
	const frame = drawingFrame(drawing);
	if (!frame) return;
	const cx = String(extent.widthEmu);
	const cy = String(extent.heightEmu);

	const frameExtent = frame.findChild("wp:extent");
	frameExtent?.setAttribute("cx", cx);
	frameExtent?.setAttribute("cy", cy);

	const shapeExtent = frame
		.findDescendant("pic:spPr")
		?.findChild("a:xfrm")
		?.findChild("a:ext");
	shapeExtent?.setAttribute("cx", cx);
	shapeExtent?.setAttribute("cy", cy);

	const blipFill = frame.findDescendant("pic:blipFill");
	if (blipFill) {
		blipFill.children = blipFill.children.filter(
			(child) => child.tag !== "a:srcRect",
		);
	}

	for (const wrapTag of ["wp:wrapTight", "wp:wrapThrough"]) {
		const polygon = frame.findChild(wrapTag)?.findChild("wp:wrapPolygon");
		if (polygon) polygon.children = fullFramePolygon();
	}
}

/** The `<wp:inline>` / `<wp:anchor>` child of a `<w:drawing>` — the element
 *  that carries the displayed size. */
function drawingFrame(drawing: XmlNode): XmlNode | undefined {
	return drawing.children.find(
		(child) => child.tag === "wp:inline" || child.tag === "wp:anchor",
	);
}

function currentExtent(drawing: XmlNode): {
	widthEmu: number;
	heightEmu: number;
} {
	const extent = drawingFrame(drawing)?.findChild("wp:extent");
	const widthEmu = Number(extent?.getAttribute("cx") ?? "0");
	const heightEmu = Number(extent?.getAttribute("cy") ?? "0");
	return {
		widthEmu: Number.isFinite(widthEmu) ? widthEmu : 0,
		heightEmu: Number.isFinite(heightEmu) ? heightEmu : 0,
	};
}

/** A wrap polygon around the whole frame. Wrap-polygon coordinates are in a
 *  fixed 21600×21600 space relative to the extent, so this rectangle stays
 *  correct whatever size the frame becomes. */
function fullFramePolygon(): XmlNode[] {
	const point = (tag: string, x: number, y: number) =>
		XmlNode.element(tag, { x: String(x), y: String(y) });
	return [
		point("wp:start", 0, 0),
		point("wp:lineTo", 0, WRAP_POLYGON_SPAN),
		point("wp:lineTo", WRAP_POLYGON_SPAN, WRAP_POLYGON_SPAN),
		point("wp:lineTo", WRAP_POLYGON_SPAN, 0),
		point("wp:lineTo", 0, 0),
	];
}

const WRAP_POLYGON_SPAN = 21600;

export type ReplacementSizing = {
	widthInches?: number;
	heightInches?: number;
	keepSize?: boolean;
};

type ReplacementExtent = {
	widthEmu: number;
	heightEmu: number;
	/** False when the new image's pixel size is unknown and the old box was
	 *  kept as-is (it may not match the new image's shape). */
	sizeKnown: boolean;
};
