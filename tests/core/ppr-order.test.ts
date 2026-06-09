import { describe, expect, test } from "bun:test";
import {
	applyParagraphOptionsInPlace,
	insertPprChildInOrder,
} from "../../src/core/blocks";
import { XmlNode } from "../../src/core/parser";

/**
 * CT_PPr child order (ECMA-376 §17.3.1.26): Word rejects a `<w:pPr>` whose
 * children are out of sequence — the classic break is `<w:jc>` (from
 * `--alignment`) appended AFTER the trailing paragraph-mark `<w:rPr>`, which
 * Word treats as "unreadable content / repair." These guard the two splice
 * sites that add to an already-built pPr (which usually ends in that rPr).
 */

function pprWithTrailingRpr(): XmlNode {
	const [pPr] = XmlNode.parse(
		'<w:pPr><w:tabs><w:tab w:val="left" w:pos="9489"/></w:tabs>' +
			'<w:ind w:left="119"/><w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr></w:pPr>',
	);
	if (!pPr) throw new Error("parse failed");
	return pPr;
}

const childTags = (pPr: XmlNode): string[] => pPr.children.map((c) => c.tag);

describe("insertPprChildInOrder", () => {
	test("splices <w:jc> before the trailing paragraph-mark <w:rPr>", () => {
		const pPr = pprWithTrailingRpr();
		insertPprChildInOrder(pPr, new XmlNode("w:jc", { "w:val": "center" }));
		const tags = childTags(pPr);
		expect(tags).toEqual(["w:tabs", "w:ind", "w:jc", "w:rPr"]);
		expect(tags.indexOf("w:jc")).toBeLessThan(tags.indexOf("w:rPr"));
	});

	test("splices <w:pStyle> to the front (it's first in CT_PPr)", () => {
		const pPr = pprWithTrailingRpr();
		insertPprChildInOrder(
			pPr,
			new XmlNode("w:pStyle", { "w:val": "Heading1" }),
		);
		expect(childTags(pPr)).toEqual(["w:pStyle", "w:tabs", "w:ind", "w:rPr"]);
	});

	test("splices <w:spacing> between <w:tabs> and <w:ind>", () => {
		const pPr = pprWithTrailingRpr();
		insertPprChildInOrder(pPr, new XmlNode("w:spacing", { "w:before": "11" }));
		expect(childTags(pPr)).toEqual(["w:tabs", "w:spacing", "w:ind", "w:rPr"]);
	});

	test("an unknown tag lands just before <w:rPr>, never after", () => {
		const pPr = pprWithTrailingRpr();
		insertPprChildInOrder(pPr, new XmlNode("w:somethingNew"));
		const tags = childTags(pPr);
		expect(tags.indexOf("w:somethingNew")).toBeLessThan(tags.indexOf("w:rPr"));
	});
});

describe("applyParagraphOptionsInPlace keeps CT_PPr order", () => {
	test("--alignment on a pPr ending in <w:rPr> puts <w:jc> before it", () => {
		const pPr = pprWithTrailingRpr();
		const children = [pPr];
		applyParagraphOptionsInPlace(children, { alignment: "center" });
		const tags = childTags(pPr);
		expect(tags.indexOf("w:jc")).toBeGreaterThanOrEqual(0);
		expect(tags.indexOf("w:jc")).toBeLessThan(tags.indexOf("w:rPr"));
	});
});
