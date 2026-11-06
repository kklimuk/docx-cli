import type { XmlNode } from "../parser";

/** Apply an `images set` change to one picture's `<w:drawing>` IN PLACE: resize
 *  its frame (`<wp:extent>` + the picture's own `<a:xfrm><a:ext>`, which must
 *  agree or Word re-derives one from the other) and/or rewrite its alt text
 *  (`<wp:docPr>`/`<pic:cNvPr>` `descr`). Works on `<wp:inline>` and
 *  `<wp:anchor>` frames alike; every other child (wrap, position, crop, effects)
 *  is left byte-identical. Returns the before/after state so the caller can
 *  report, audit, and detect a no-op. Throws `ImageFrameError` when the drawing
 *  isn't a single picture we can resize honestly. */
export function setImageFrame(
	drawing: XmlNode,
	change: ImageFrameChange,
): ImageFrameResult {
	const frame = drawingFrame(drawing);
	if (!frame) {
		throw new ImageFrameError(
			"UNSUPPORTED",
			"This drawing has no <wp:inline>/<wp:anchor> frame to resize.",
		);
	}
	const before = readFrameState(frame);
	const after: ImageFrameState = { ...before };

	if (hasSizeChange(change)) {
		if (!isSinglePicture(frame)) {
			throw new ImageFrameError(
				"UNSUPPORTED",
				"This image is part of a drawing group or shape, not a lone picture, so it can't be resized on its own.",
				"Resize the group in Word, or replace it with `docx images replace`.",
			);
		}
		const size = resolveFrameSize(before, change);
		after.widthEmu = size.widthEmu;
		after.heightEmu = size.heightEmu;
		writeFrameSize(frame, size);
	}

	if (change.alt !== undefined) {
		after.alt = change.alt;
		writeFrameAlt(frame, change.alt);
	}

	return {
		before,
		after,
		resized:
			after.widthEmu !== before.widthEmu ||
			after.heightEmu !== before.heightEmu,
		altChanged: after.alt !== before.alt,
	};
}

/** The `<wp:inline>` (in-flow) or `<wp:anchor>` (floating) frame — the one
 *  element child of `<w:drawing>` that carries the extent. */
function drawingFrame(drawing: XmlNode): XmlNode | undefined {
	return drawing.children.find(
		(child) => child.tag === "wp:inline" || child.tag === "wp:anchor",
	);
}

function readFrameState(frame: XmlNode): ImageFrameState {
	const extent = frame.findChild("wp:extent");
	return {
		widthEmu: Number(extent?.getAttribute("cx") ?? "0") || 0,
		heightEmu: Number(extent?.getAttribute("cy") ?? "0") || 0,
		alt: frame.findChild("wp:docPr")?.getAttribute("descr") ?? "",
	};
}

function hasSizeChange(change: ImageFrameChange): boolean {
	return (
		change.widthEmu !== undefined ||
		change.heightEmu !== undefined ||
		change.fitWidthEmu !== undefined
	);
}

/** A resizable lone picture: `<a:graphicData>` holds a `pic:pic`, not a
 *  `wpg:wgp` group or `wps:wsp` shape whose children carry their own geometry
 *  (scaling only the outer extent would distort or clip them). */
function isSinglePicture(frame: XmlNode): boolean {
	const graphicData = frame.findChild("a:graphic")?.findChild("a:graphicData");
	return graphicData?.getAttribute("uri") === DRAWINGML_PICTURE_URI;
}

/** The new extent. A single axis keeps the frame's CURRENT aspect ratio (Word's
 *  lock-aspect corner drag); both axes are honored as given; `fitWidthEmu`
 *  shrinks a wider frame to that width, aspect kept, and never enlarges. */
function resolveFrameSize(
	current: ImageFrameState,
	change: ImageFrameChange,
): { widthEmu: number; heightEmu: number } {
	const { widthEmu, heightEmu, fitWidthEmu } = change;
	if (widthEmu !== undefined && heightEmu !== undefined) {
		return { widthEmu, heightEmu };
	}
	const hasAspect = current.widthEmu > 0 && current.heightEmu > 0;
	if ((widthEmu !== undefined || heightEmu !== undefined) && !hasAspect) {
		throw new ImageFrameError(
			"USAGE",
			"The image has no current size to keep the aspect ratio from.",
			"Pass both --width and --height.",
		);
	}
	if (widthEmu !== undefined) {
		return {
			widthEmu,
			heightEmu: Math.round((widthEmu * current.heightEmu) / current.widthEmu),
		};
	}
	if (heightEmu !== undefined) {
		return {
			widthEmu: Math.round((heightEmu * current.widthEmu) / current.heightEmu),
			heightEmu,
		};
	}
	if (
		fitWidthEmu !== undefined &&
		hasAspect &&
		current.widthEmu > fitWidthEmu
	) {
		return {
			widthEmu: fitWidthEmu,
			heightEmu: Math.round(
				(fitWidthEmu * current.heightEmu) / current.widthEmu,
			),
		};
	}
	return { widthEmu: current.widthEmu, heightEmu: current.heightEmu };
}

/** Write the size to BOTH places Word keeps it: the frame's `<wp:extent>` and
 *  the picture's `<pic:spPr><a:xfrm><a:ext>` (looked up by path — a bare
 *  `findDescendant("a:ext")` would hit an `<a:extLst><a:ext>` extension). */
function writeFrameSize(
	frame: XmlNode,
	size: { widthEmu: number; heightEmu: number },
): void {
	const extent = frame.findChild("wp:extent");
	extent?.setAttribute("cx", String(size.widthEmu));
	extent?.setAttribute("cy", String(size.heightEmu));
	const transformExtent = frame
		.findDescendant("pic:spPr")
		?.findChild("a:xfrm")
		?.findChild("a:ext");
	transformExtent?.setAttribute("cx", String(size.widthEmu));
	transformExtent?.setAttribute("cy", String(size.heightEmu));
}

/** Alt text lives on `<wp:docPr descr>` (what Word's Alt Text pane edits and
 *  what `read` shows as `![alt]`), mirrored onto `<pic:cNvPr descr>` as Word
 *  writes it. An empty string clears it. */
function writeFrameAlt(frame: XmlNode, alt: string): void {
	const carriers = [
		frame.findChild("wp:docPr"),
		frame.findDescendant("pic:cNvPr"),
	];
	for (const carrier of carriers) {
		if (!carrier) continue;
		if (alt === "") delete carrier.attributes.descr;
		else carrier.setAttribute("descr", alt);
	}
}

export const DRAWINGML_PICTURE_URI =
	"http://schemas.openxmlformats.org/drawingml/2006/picture";

export class ImageFrameError extends Error {
	constructor(
		public code: "USAGE" | "UNSUPPORTED" | "IMAGE_NOT_FOUND",
		message: string,
		public hint?: string,
	) {
		super(message);
	}
}

/** What `images set` asks for. Sizes are EMU; omit a field to leave it alone. */
export type ImageFrameChange = {
	widthEmu?: number;
	heightEmu?: number;
	/** Shrink to at most this width (aspect kept); never enlarges. */
	fitWidthEmu?: number;
	/** New alt text; `""` clears it. */
	alt?: string;
};

type ImageFrameState = { widthEmu: number; heightEmu: number; alt: string };

export type ImageFrameResult = {
	before: ImageFrameState;
	after: ImageFrameState;
	resized: boolean;
	altChanged: boolean;
};
