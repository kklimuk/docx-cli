import { describe, expect, test } from "bun:test";
import { runTextLength, sliceRun, XmlNode } from "@core/parser";

function parseRun(xml: string): XmlNode {
	const tree = XmlNode.parse(
		`<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${xml.replace(/^<w:r[^>]*>|<\/w:r>$/g, "")}</w:r>`,
	);
	const run = XmlNode.findRoot(tree, "w:r");
	if (!run) throw new Error("could not parse run");
	return run;
}

function summarize(run: XmlNode): { tag: string; text?: string }[] {
	const out: { tag: string; text?: string }[] = [];
	for (const child of run.children) {
		if (child.tag === "w:t" || child.tag === "w:delText") {
			out.push({ tag: child.tag, text: child.collectText() });
			continue;
		}
		out.push({ tag: child.tag });
	}
	return out;
}

describe("sliceRun (Bug B)", () => {
	test("B1: tab + foobar, slice [0,3) → tab + fo (tab occupies offset 0)", () => {
		const run = parseRun(`<w:r><w:tab/><w:t>foobar</w:t></w:r>`);
		// Tab is one character wide: tab=0, f=1, o=2, o=3, … so [0,3) is tab + "fo".
		expect(summarize(sliceRun(run, 0, 3))).toEqual([
			{ tag: "w:tab" },
			{ tag: "w:t", text: "fo" },
		]);
	});

	test("B2: tab + foobar, slice [3,6) → oba (tab is before the slice)", () => {
		const run = parseRun(`<w:r><w:tab/><w:t>foobar</w:t></w:r>`);
		// tab=0, foobar=1..6, so [3,6) is "oba".
		expect(summarize(sliceRun(run, 3, 6))).toEqual([
			{ tag: "w:t", text: "oba" },
		]);
	});

	test("B3: foo + tab + bar, slice [3,6) → tab + ba (tab at offset 3 belongs to this slice)", () => {
		const run = parseRun(`<w:r><w:t>foo</w:t><w:tab/><w:t>bar</w:t></w:r>`);
		// foo=0..2, tab=3, bar=4..6, so [3,6) is tab + "ba".
		expect(summarize(sliceRun(run, 3, 6))).toEqual([
			{ tag: "w:tab" },
			{ tag: "w:t", text: "ba" },
		]);
	});

	test("B4: foo + tab + bar, slice [0,3) → foo only (tab at offset 3 not in [0,3))", () => {
		const run = parseRun(`<w:r><w:t>foo</w:t><w:tab/><w:t>bar</w:t></w:r>`);
		expect(summarize(sliceRun(run, 0, 3))).toEqual([
			{ tag: "w:t", text: "foo" },
		]);
	});

	test("B5: <w:rPr> is preserved in every non-empty slice", () => {
		const run = parseRun(
			`<w:r><w:rPr><w:b/></w:rPr><w:t>foo</w:t><w:tab/><w:t>bar</w:t></w:r>`,
		);
		const pre = sliceRun(run, 0, 3);
		const cut = sliceRun(run, 3, 6);
		expect(pre.findChild("w:rPr")?.findChild("w:b")).toBeDefined();
		expect(cut.findChild("w:rPr")?.findChild("w:b")).toBeDefined();
	});

	test("partition is exhaustive: pre + cut + post recovers every non-rPr child exactly once", () => {
		const run = parseRun(`<w:r><w:t>foo</w:t><w:tab/><w:t>bar</w:t></w:r>`);
		// Total width is 7 (foo=3, tab=1, bar=3): foo=0..2, tab=3, bar=4..6.
		const pre = summarize(sliceRun(run, 0, 2));
		const cut = summarize(sliceRun(run, 2, 5));
		const post = summarize(sliceRun(run, 5, 7));
		const combined = [...pre, ...cut, ...post];
		expect(combined).toEqual([
			{ tag: "w:t", text: "fo" },
			{ tag: "w:t", text: "o" },
			{ tag: "w:tab" },
			{ tag: "w:t", text: "b" },
			{ tag: "w:t", text: "ar" },
		]);
	});

	test("delText slices like w:t and the slice keeps the delText tag", () => {
		const run = parseRun(`<w:r><w:delText>foobar</w:delText></w:r>`);
		expect(summarize(sliceRun(run, 0, 3))).toEqual([
			{ tag: "w:delText", text: "foo" },
		]);
	});
});

describe("single-character markers (noBreakHyphen / softHyphen / sym)", () => {
	test("runTextLength counts noBreakHyphen, softHyphen, and sym as one char each", () => {
		const run = parseRun(
			`<w:r><w:t>co</w:t><w:noBreakHyphen/><w:t>op</w:t></w:r>`,
		);
		// "co" (2) + nbhyphen (1) + "op" (2) = 5
		expect(runTextLength(run)).toBe(5);

		const symRun = parseRun(
			`<w:r><w:t>x</w:t><w:sym w:font="Symbol" w:char="44"/><w:t>y</w:t></w:r>`,
		);
		// "x" (1) + sym (1) + "y" (1) = 3
		expect(runTextLength(symRun)).toBe(3);

		const softRun = parseRun(
			`<w:r><w:t>foo</w:t><w:softHyphen/><w:t>bar</w:t></w:r>`,
		);
		expect(runTextLength(softRun)).toBe(7);
	});

	test("sliceRun: noBreakHyphen at offset 2 is owned by the slice [2, 3)", () => {
		const run = parseRun(
			`<w:r><w:t>co</w:t><w:noBreakHyphen/><w:t>op</w:t></w:r>`,
		);
		// Pre [0, 2): "co" only
		expect(summarize(sliceRun(run, 0, 2))).toEqual([
			{ tag: "w:t", text: "co" },
		]);
		// Cut [2, 3): just the noBreakHyphen
		expect(summarize(sliceRun(run, 2, 3))).toEqual([
			{ tag: "w:noBreakHyphen" },
		]);
		// Post [3, 5): "op"
		expect(summarize(sliceRun(run, 3, 5))).toEqual([
			{ tag: "w:t", text: "op" },
		]);
	});

	test("sliceRun: noBreakHyphen + tab survive in the right slices simultaneously", () => {
		const run = parseRun(
			`<w:r><w:t>a</w:t><w:noBreakHyphen/><w:tab/><w:t>b</w:t></w:r>`,
		);
		// Offsets: 'a'=0, nbhyphen=1, tab=2, 'b'=3 (tab is one character wide).
		// Slice [0, 2): "a" + nbhyphen (offset 1 in [0, 2))
		expect(summarize(sliceRun(run, 0, 2))).toEqual([
			{ tag: "w:t", text: "a" },
			{ tag: "w:noBreakHyphen" },
		]);
		// Slice [2, 3): the tab alone (offset 2)
		expect(summarize(sliceRun(run, 2, 3))).toEqual([{ tag: "w:tab" }]);
		// Slice [3, 4): 'b'
		expect(summarize(sliceRun(run, 3, 4))).toEqual([{ tag: "w:t", text: "b" }]);
	});

	test("runTextLength counts tab, line break, and cr as one char; page break as zero", () => {
		expect(
			runTextLength(parseRun(`<w:r><w:t>a</w:t><w:tab/><w:t>b</w:t></w:r>`)),
		).toBe(3);
		expect(
			runTextLength(parseRun(`<w:r><w:t>a</w:t><w:br/><w:t>b</w:t></w:r>`)),
		).toBe(3);
		expect(
			runTextLength(parseRun(`<w:r><w:t>a</w:t><w:cr/><w:t>b</w:t></w:r>`)),
		).toBe(3);
		// A page/column break renders as nothing in read, so it stays zero-width.
		expect(
			runTextLength(
				parseRun(`<w:r><w:t>a</w:t><w:br w:type="page"/><w:t>b</w:t></w:r>`),
			),
		).toBe(2);
	});
});
