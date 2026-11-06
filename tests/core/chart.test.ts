import { describe, expect, test } from "bun:test";
import { Document, type NativeChartRun, type Paragraph } from "@core";
import { Charts, chartToCsv } from "@core/chart";

const CHART_NS =
	'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function drawing(rId: string, { anchor = false } = {}): string {
	const placement = anchor
		? '<wp:anchor><wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH><wp:extent cx="914400" cy="457200"/><wp:wrapSquare wrapText="bothSides"/>'
		: '<wp:inline><wp:extent cx="914400" cy="457200"/>';
	const close = anchor ? "</wp:anchor>" : "</wp:inline>";
	return `<w:r><w:drawing>${placement}<wp:docPr id="1" name="Chart 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart ${CHART_NS} r:id="${rId}"/></a:graphicData></a:graphic>${close}</w:drawing></w:r>`;
}

function chartDocument(
	bodyXml: string,
	charts: Record<string, { xml: string; relsXml?: string }>,
): Document {
	const relationships = Object.keys(charts)
		.map(
			(partName, index) =>
				`<Relationship Id="rIdC${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="${partName.replace("word/", "")}"/>`,
		)
		.join("");
	return Document.fromXml({
		documentXml: `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyXml}<w:sectPr/></w:body></w:document>`,
		relationshipsXml: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`,
		chartParts: charts,
	});
}

function chartSpace(inner: string, trailing = ""): string {
	return `<c:chartSpace ${CHART_NS}><c:chart>${inner}</c:chart>${trailing}</c:chartSpace>`;
}

function onlyChart(document: Document): NativeChartRun {
	const charts = new Charts(document).list();
	expect(charts).toHaveLength(1);
	const run = new Charts(document).get("chart0");
	if (!run) throw new Error("expected chart0");
	return run;
}

const BAR_SERIES = `<c:ser><c:idx val="0"/><c:order val="0"/>
<c:tx><c:strRef><c:f>'Summary, FY'!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
<c:cat><c:strRef><c:f>'Summary, FY'!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>North</c:v></c:pt><c:pt idx="1"><c:v>South, "Gulf"</c:v></c:pt><c:pt idx="2"><c:v>West</c:v></c:pt></c:strCache></c:strRef></c:cat>
<c:val><c:numRef><c:f>'Summary, FY'!$B$2:$B$4</c:f><c:numCache><c:formatCode>0.0%</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>0.25</c:v></c:pt><c:pt idx="2"><c:v>0.5</c:v></c:pt></c:numCache></c:numRef></c:val>
</c:ser>`.replace(/\n/g, "");

describe("chart reader", () => {
	test("a column chart: id, type, title, series from the caches, gap as null", () => {
		const document = chartDocument(`<w:p>${drawing("rIdC0")}</w:p>`, {
			"word/charts/chart1.xml": {
				xml: chartSpace(
					`<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Revenue </a:t></a:r><a:r><a:t>by Region</a:t></a:r></a:p><a:p><a:r><a:t>FY2026</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/>${BAR_SERIES}</c:barChart><c:catAx/><c:valAx/></c:plotArea>`,
				),
			},
		});
		const chart = onlyChart(document);
		expect(chart).toMatchObject({
			type: "chart",
			kind: "chart",
			id: "chart0",
			chartTypes: ["column"],
			title: "Revenue by Region FY2026",
			dataSource: "none",
			widthEmu: 914400,
			heightEmu: 457200,
		});
		expect(chart.series).toEqual([
			{
				name: "Revenue",
				nameRef: "'Summary, FY'!$B$1",
				categories: ["North", 'South, "Gulf"', "West"],
				categoriesRef: "'Summary, FY'!$A$2:$A$4",
				values: [0.25, null, 0.5],
				valuesRef: "'Summary, FY'!$B$2:$B$4",
				formatCode: "0.0%",
			},
		]);
	});

	test("CSV: header + one row per point, RFC 4180 quoting, a gap is an empty cell", () => {
		const document = chartDocument(`<w:p>${drawing("rIdC0")}</w:p>`, {
			"word/charts/chart1.xml": {
				xml: chartSpace(
					`<c:plotArea><c:barChart><c:barDir val="bar"/>${BAR_SERIES}</c:barChart></c:plotArea>`,
				),
			},
		});
		const chart = onlyChart(document);
		expect(chart.chartTypes).toEqual(["bar"]);
		expect(chartToCsv(chart)).toBe(
			'category,Revenue\nNorth,0.25\n"South, ""Gulf""",\nWest,0.5\n',
		);
	});

	test("a combo chart lists every chart group; literals and unnamed series read too", () => {
		const document = chartDocument(`<w:p>${drawing("rIdC0")}</w:p>`, {
			"word/charts/chart1.xml": {
				xml: chartSpace(
					`<c:plotArea><c:bar3DChart><c:barDir val="col"/>${BAR_SERIES}</c:bar3DChart><c:lineChart><c:ser><c:idx val="1"/><c:order val="1"/><c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>a</c:v></c:pt><c:pt idx="1"><c:v>b</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart></c:plotArea>`,
				),
			},
		});
		const chart = onlyChart(document);
		expect(chart.chartTypes).toEqual(["column3D", "line"]);
		expect(chart.series[1]).toEqual({ categories: ["a", "b"], values: [1, 2] });
		expect(new Charts(document).list()[0]).toMatchObject({
			type: "column3D+line",
			series: ["Revenue", "Series 2"],
			points: 3,
		});
	});

	test("multi-level categories read the innermost level", () => {
		const document = chartDocument(`<w:p>${drawing("rIdC0")}</w:p>`, {
			"word/charts/chart1.xml": {
				xml: chartSpace(
					`<c:plotArea><c:pieChart><c:ser><c:cat><c:multiLvlStrRef><c:f>Sheet1!$A$2:$B$3</c:f><c:multiLvlStrCache><c:ptCount val="2"/><c:lvl><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt></c:lvl><c:lvl><c:pt idx="0"><c:v>2026</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat><c:val><c:numRef><c:f>Sheet1!$C$2:$C$3</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:pieChart></c:plotArea>`,
				),
			},
		});
		const chart = onlyChart(document);
		expect(chart.chartTypes).toEqual(["pie"]);
		expect(chart.series[0]?.categories).toEqual(["Jan", "Feb"]);
		expect(chart.series[0]?.categoriesRef).toBe("Sheet1!$A$2:$B$3");
	});

	test("data source: an embedded workbook resolves against the chart part; a link is external", () => {
		const relsXml = (target: string, mode = "") =>
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="${target}"${mode}/></Relationships>`;
		const space = chartSpace(
			"<c:plotArea><c:areaChart/></c:plotArea>",
			'<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>',
		);
		const document = chartDocument(
			`<w:p>${drawing("rIdC0")}</w:p><w:p>${drawing("rIdC1", { anchor: true })}</w:p>`,
			{
				"word/charts/chart1.xml": {
					xml: space,
					relsXml: relsXml("../embeddings/Microsoft_Excel_Worksheet1.xlsx"),
				},
				"word/charts/chart2.xml": {
					xml: space,
					relsXml: relsXml("file:///C:/Book.xlsx", ' TargetMode="External"'),
				},
			},
		);
		const charts = new Charts(document);
		expect(charts.get("chart0")).toMatchObject({
			chartTypes: ["area"],
			dataSource: "embedded",
			workbook: "word/embeddings/Microsoft_Excel_Worksheet1.xlsx",
		});
		expect(charts.get("chart1")).toMatchObject({
			dataSource: "external",
			floating: true,
			wrap: "square",
			align: "center",
		});
		expect(charts.get("chart1")?.workbook).toBeUndefined();
	});

	test("chartN is positional across body and cells; a dangling chart still gets its id", () => {
		const document = chartDocument(
			`<w:p>${drawing("rIdMissing")}</w:p><w:tbl><w:tr><w:tc><w:p>${drawing("rIdC0")}</w:p></w:tc></w:tr></w:tbl>`,
			{
				"word/charts/chart1.xml": {
					xml: chartSpace("<c:plotArea><c:lineChart/></c:plotArea>"),
				},
			},
		);
		const listing = new Charts(document).list();
		expect(listing.map(({ id, at, type }) => ({ id, at, type }))).toEqual([
			{ id: "chart0", at: "p0", type: "unknown" },
			{ id: "chart1", at: "t0:r0c0:p0", type: "line" },
		]);
	});

	test("an out-of-range point index or count is ignored, not allocated", () => {
		const document = chartDocument(`<w:p>${drawing("rIdC0")}</w:p>`, {
			"word/charts/chart1.xml": {
				xml: chartSpace(
					'<c:plotArea><c:lineChart><c:ser><c:val><c:numLit><c:ptCount val="999999999"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="999999999"><c:v>2</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart></c:plotArea>',
				),
			},
		});
		expect(onlyChart(document).series[0]?.values).toEqual([1]);
	});

	test("SmartArt and shapes stay bare placeholders with no chart id", () => {
		const document = chartDocument(
			'<w:p><w:r><w:drawing><dgm:relIds r:dm="rId9"/></w:drawing></w:r></w:p>',
			{},
		);
		const paragraph = document.body.blocks[0] as Paragraph;
		expect(paragraph.runs).toEqual([{ type: "chart", kind: "smartart" }]);
		expect(new Charts(document).list()).toEqual([]);
	});
});
