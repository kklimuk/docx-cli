import type { XmlNode } from "../parser";
import type { CellMargins } from "./types";

/** Cell padding from a `<w:tblCellMar>`/`<w:tcMar>` element, in twips. Reads the
 * transitional `left`/`right` and the strict `start`/`end` alike; `nil` is 0 and
 * a percentage side (not meaningful for padding) is skipped. Undefined when the
 * element is absent or declares no readable side. Shared by the AST reader (a
 * table's/cell's own margins) and `StylesView` (a table style's). */
export function readCellMargins(
	node: XmlNode | undefined,
): CellMargins | undefined {
	if (!node) return undefined;
	const margins: CellMargins = {};
	for (const [side, tag] of CELL_MARGIN_TAGS) {
		if (margins[side] !== undefined) continue;
		const edge = node.findChild(tag);
		if (!edge) continue;
		const type = edge.getAttribute("w:type") ?? "dxa";
		if (type === "nil") margins[side] = 0;
		if (type !== "dxa") continue;
		const value = Number(edge.getAttribute("w:w"));
		if (Number.isFinite(value)) margins[side] = value;
	}
	return Object.keys(margins).length > 0 ? margins : undefined;
}

/** The four padded sides in CSS order (top, right, bottom, left) — the order
 * the `--cell-margins` flag takes and the read hints print. */
export const CELL_MARGIN_SIDES = ["top", "right", "bottom", "left"] as const;

/** A margins element's children in CT_TcMar / CT_TblCellMar schema order, each
 * paired with the side it pads: the strict `start`/`end` and the transitional
 * `left`/`right` both pad left/right. */
export const CELL_MARGIN_TAGS = [
	["top", "w:top"],
	["left", "w:start"],
	["left", "w:left"],
	["bottom", "w:bottom"],
	["right", "w:end"],
	["right", "w:right"],
] as const;
