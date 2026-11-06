import { describe, expect, test } from "bun:test";
import { XmlNode } from "@core";
import { applyReplacementExtent, fitReplacementExtent } from "@core/image";

const EMU_PER_INCH = 914400;
const LETTER_TEXT_AREA = {
	widthEmu: 6.5 * EMU_PER_INCH,
	heightEmu: 9 * EMU_PER_INCH,
};

function drawing(frame: string): XmlNode {
	const [node] = XmlNode.parse(`<w:drawing>${frame}</w:drawing>`);
	if (!node) throw new Error("parse failed");
	return node;
}

const WIDE_INLINE = `<wp:inline><wp:extent cx="${4 * EMU_PER_INCH}" cy="${EMU_PER_INCH}"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId9"/><a:srcRect l="10000" r="5000"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${4 * EMU_PER_INCH}" cy="${EMU_PER_INCH}"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline>`;

describe("fitReplacementExtent", () => {
	const tall = { pixelWidth: 100, pixelHeight: 400 };

	test("default keeps the displayed width and recomputes the height", () => {
		const fit = fitReplacementExtent(
			drawing(
				WIDE_INLINE.replaceAll(
					String(4 * EMU_PER_INCH),
					String(2 * EMU_PER_INCH),
				),
			),
			tall,
			{},
			LETTER_TEXT_AREA,
		);
		expect(fit).toEqual({
			widthEmu: 2 * EMU_PER_INCH,
			heightEmu: 8 * EMU_PER_INCH,
			sizeKnown: true,
		});
	});

	test("default scales down (aspect kept) to fit the text area", () => {
		// 4in wide at 1:4 would be 16in tall — fit to the 9in text height.
		const fit = fitReplacementExtent(
			drawing(WIDE_INLINE),
			tall,
			{},
			LETTER_TEXT_AREA,
		);
		expect(fit?.heightEmu).toBe(9 * EMU_PER_INCH);
		expect(fit?.widthEmu).toBe(2.25 * EMU_PER_INCH);
	});

	test("unknown pixel size keeps the old box and says so", () => {
		const fit = fitReplacementExtent(
			drawing(WIDE_INLINE),
			{},
			{},
			LETTER_TEXT_AREA,
		);
		expect(fit).toEqual({
			widthEmu: 4 * EMU_PER_INCH,
			heightEmu: EMU_PER_INCH,
			sizeKnown: false,
		});
	});

	test("one explicit axis needs an aspect ratio", () => {
		expect(
			fitReplacementExtent(
				drawing(WIDE_INLINE),
				{},
				{ widthInches: 1 },
				LETTER_TEXT_AREA,
			),
		).toBeNull();
		expect(
			fitReplacementExtent(
				drawing(WIDE_INLINE),
				{},
				{ widthInches: 1, heightInches: 2 },
				LETTER_TEXT_AREA,
			),
		).toEqual({
			widthEmu: EMU_PER_INCH,
			heightEmu: 2 * EMU_PER_INCH,
			sizeKnown: true,
		});
	});
});

describe("applyReplacementExtent", () => {
	test("updates both extents and drops the crop", () => {
		const node = drawing(WIDE_INLINE);
		applyReplacementExtent(node, { widthEmu: 100, heightEmu: 400 });
		const xml = XmlNode.serialize([node]);
		expect(xml).toContain('<wp:extent cx="100" cy="400"/>');
		expect(xml).toContain('<a:ext cx="100" cy="400"/>');
		expect(xml).not.toContain("a:srcRect");
		expect(xml).toContain("<a:stretch>");
	});

	test("resets a floating image's wrap polygon to the full frame", () => {
		const node = drawing(
			`<wp:anchor><wp:extent cx="10" cy="10"/><wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="5" y="5"/><wp:lineTo x="900" y="20000"/><wp:lineTo x="5" y="5"/></wp:wrapPolygon></wp:wrapTight></wp:anchor>`,
		);
		applyReplacementExtent(node, { widthEmu: 20, heightEmu: 40 });
		const xml = XmlNode.serialize([node]);
		expect(xml).toContain('<wp:extent cx="20" cy="40"/>');
		expect(xml).toContain(
			'<wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/>',
		);
	});
});
