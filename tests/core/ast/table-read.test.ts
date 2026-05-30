import { describe, expect, test } from "bun:test";
import { Document } from "@core/ast/document";
import type { Body } from "@core/ast/document/body";
import type { Table } from "@core/ast/types";

function buildSyntheticView(bodyXml: string): Body {
	return Document.fromXml({
		documentXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>${bodyXml}<w:sectPr/></w:body>
</w:document>`,
	}).body;
}

function firstTable(doc: Body): Table {
	const block = doc.blocks[0];
	if (!block || block.type !== "table") {
		throw new Error("expected table as first block");
	}
	return block;
}

describe("readTable AST fields", () => {
	test("grid widths are extracted from <w:tblGrid>", () => {
		const doc = buildSyntheticView(
			`<w:tbl>
				<w:tblGrid>
					<w:gridCol w:w="1440"/>
					<w:gridCol w:w="2880"/>
					<w:gridCol w:w="5040"/>
				</w:tblGrid>
				<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr>
			</w:tbl>`,
		);
		const table = firstTable(doc);
		expect(table.grid).toEqual([1440, 2880, 5040]);
	});

	test("table-level width is extracted from <w:tblPr><w:tblW/>", () => {
		const doc = buildSyntheticView(
			`<w:tbl>
				<w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>
				<w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>
				<w:tr><w:tc><w:p/></w:tc></w:tr>
			</w:tbl>`,
		);
		const table = firstTable(doc);
		expect(table.width).toEqual({ value: 5000, unit: "pct" });
	});

	test("missing <w:tblPr> leaves table.width undefined", () => {
		const doc = buildSyntheticView(
			`<w:tbl>
				<w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>
				<w:tr><w:tc><w:p/></w:tc></w:tr>
			</w:tbl>`,
		);
		const table = firstTable(doc);
		expect(table.width).toBeUndefined();
	});

	test("unknown w:type falls back to dxa", () => {
		const doc = buildSyntheticView(
			`<w:tbl>
				<w:tblPr><w:tblW w:w="4320" w:type="other"/></w:tblPr>
				<w:tblGrid><w:gridCol w:w="4320"/></w:tblGrid>
				<w:tr><w:tc><w:p/></w:tc></w:tr>
			</w:tbl>`,
		);
		const table = firstTable(doc);
		expect(table.width).toEqual({ value: 4320, unit: "dxa" });
	});

	test("gridSpan > 1 is captured on the cell", () => {
		const doc = buildSyntheticView(
			`<w:tbl>
				<w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/></w:tblGrid>
				<w:tr>
					<w:tc>
						<w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
						<w:p/>
					</w:tc>
				</w:tr>
			</w:tbl>`,
		);
		const table = firstTable(doc);
		expect(table.rows[0]?.cells[0]?.gridSpan).toBe(2);
	});

	test("gridSpan of 1 is omitted (default)", () => {
		const doc = buildSyntheticView(
			`<w:tbl>
				<w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
				<w:tr>
					<w:tc>
						<w:tcPr><w:gridSpan w:val="1"/></w:tcPr>
						<w:p/>
					</w:tc>
				</w:tr>
			</w:tbl>`,
		);
		const table = firstTable(doc);
		expect(table.rows[0]?.cells[0]?.gridSpan).toBeUndefined();
	});

	test("vMerge restart and continue are distinguished", () => {
		const doc = buildSyntheticView(
			`<w:tbl>
				<w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
				<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p/></w:tc></w:tr>
				<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr>
				<w:tr><w:tc><w:tcPr><w:vMerge w:val="continue"/></w:tcPr><w:p/></w:tc></w:tr>
			</w:tbl>`,
		);
		const table = firstTable(doc);
		expect(table.rows[0]?.cells[0]?.vMerge).toBe("restart");
		expect(table.rows[1]?.cells[0]?.vMerge).toBe("continue");
		expect(table.rows[2]?.cells[0]?.vMerge).toBe("continue");
	});

	test("cell-level <w:tcW> populates cell.width", () => {
		const doc = buildSyntheticView(
			`<w:tbl>
				<w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
				<w:tr>
					<w:tc>
						<w:tcPr><w:tcW w:w="3600" w:type="dxa"/></w:tcPr>
						<w:p/>
					</w:tc>
				</w:tr>
			</w:tbl>`,
		);
		const table = firstTable(doc);
		expect(table.rows[0]?.cells[0]?.width).toEqual({
			value: 3600,
			unit: "dxa",
		});
	});

	test("cells without <w:tcPr> stay clean (no merge / width fields)", () => {
		const doc = buildSyntheticView(
			`<w:tbl>
				<w:tblGrid><w:gridCol w:w="2880"/></w:tblGrid>
				<w:tr><w:tc><w:p/></w:tc></w:tr>
			</w:tbl>`,
		);
		const cell = firstTable(doc).rows[0]?.cells[0];
		expect(cell?.gridSpan).toBeUndefined();
		expect(cell?.vMerge).toBeUndefined();
		expect(cell?.width).toBeUndefined();
	});

	test("table without <w:tblGrid> falls back to empty grid", () => {
		const doc = buildSyntheticView(
			`<w:tbl>
				<w:tr><w:tc><w:p/></w:tc></w:tr>
			</w:tbl>`,
		);
		const table = firstTable(doc);
		expect(table.grid).toEqual([]);
	});
});
