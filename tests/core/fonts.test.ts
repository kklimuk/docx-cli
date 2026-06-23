import { describe, expect, test } from "bun:test";
import { Document, Fonts, XmlNode } from "@core";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** A synthetic doc whose styles part pins an explicit Heading1 font — and, via
 *  `Pkg.empty`, has NO theme part, exercising the graceful no-theme branch. */
function docWithExplicitHeading(): Document {
	return Document.fromXml({
		documentXml: `<w:document xmlns:w="${W_NS}"><w:body><w:p/><w:sectPr/></w:body></w:document>`,
		stylesXml: `<w:styles xmlns:w="${W_NS}"><w:style w:type="paragraph" w:styleId="Heading1"><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/></w:rPr></w:style></w:styles>`,
	});
}

describe("Fonts.setDefault (core)", () => {
	test("sets docDefaults, reports explicit styles, no-theme → themeUpdated false", async () => {
		const doc = docWithExplicitHeading();
		const result = await new Fonts(doc).setDefault("Times New Roman");

		expect(result.themeUpdated).toBe(false); // Pkg.empty has no theme part
		expect(result.explicitStyles).toContain("Heading1");
		expect(result.repointed).toBe(0); // default scope leaves explicit fonts alone

		const styles = XmlNode.serialize(doc.ensureStyles().tree);
		expect(styles).toMatch(/<w:docDefaults>[\s\S]*w:ascii="Times New Roman"/);
		// Heading1's own font is untouched in the default scope.
		expect(styles).toMatch(/Heading1[\s\S]*?w:ascii="Calibri Light"/);
	});

	test("--all repoints the explicit style font", async () => {
		const doc = docWithExplicitHeading();
		const result = await new Fonts(doc).setDefault("Georgia", { all: true });

		expect(result.repointed).toBeGreaterThan(0);
		const styles = XmlNode.serialize(doc.ensureStyles().tree);
		expect(styles).toMatch(/Heading1[\s\S]*?w:ascii="Georgia"/);
		expect(styles).not.toContain("Calibri Light");
	});

	test("--size writes docDefaults sz/szCs in half-points", async () => {
		const doc = docWithExplicitHeading();
		await new Fonts(doc).setDefault("Calibri", { sizeHalfPoints: 24 });
		const styles = XmlNode.serialize(doc.ensureStyles().tree);
		expect(styles).toMatch(/<w:docDefaults>[\s\S]*<w:sz w:val="24"/);
		expect(styles).toMatch(/<w:docDefaults>[\s\S]*<w:szCs w:val="24"/);
	});

	// Regression: --size must insert <w:sz>/<w:szCs> at the canonical CT_RPr slot
	// (after color/kern, before highlight/lang), NOT merely after <w:rFonts>.
	// A docDefaults rPr that already carries b/color/kern but no sz is valid OOXML
	// (Pandoc / Google-Docs / LibreOffice exports produce it); the old
	// "insert right after rFonts" put sz before those → Word rejects the file.
	test("--size keeps CT_RPr order when other run props precede sz's slot", async () => {
		const doc = Document.fromXml({
			documentXml: `<w:document xmlns:w="${W_NS}"><w:body><w:p/><w:sectPr/></w:body></w:document>`,
			stylesXml: `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:b/><w:color w:val="FF0000"/><w:kern w:val="2"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
		});
		await new Fonts(doc).setDefault("Georgia", { sizeHalfPoints: 24 });
		const xml = XmlNode.serialize(doc.ensureStyles().tree);

		const at = (re: RegExp): number => xml.search(re);
		// color/kern come BEFORE sz; sz before szCs; szCs before nothing later here.
		expect(at(/<w:color/)).toBeLessThan(at(/<w:sz /));
		expect(at(/<w:kern/)).toBeLessThan(at(/<w:sz /));
		expect(at(/<w:sz /)).toBeLessThan(at(/<w:szCs/));
		expect(xml).toContain('<w:sz w:val="24"');
	});
});
