import { resolve } from "node:path";
import { $ } from "bun";

// Pin core.xml timestamps so rebuilds are byte-deterministic.
process.env.DOCX_CLI_NOW ??= "2026-11-06T00:00:00Z";

/**
 * Build tests/fixtures/charts.docx — two native Word charts for the read-only
 * chart surface (`chartN`, the `docx:chart` hint, `charts list/extract`):
 *
 *   - chart0, in the body: a titled clustered COLUMN chart, two series × four
 *     quarters, with a GAP (Q3 has no East value — a missing `c:pt`);
 *   - chart1, inside a table cell: a COMBO chart (`barChart` + `lineChart`),
 *     so the `chartN` numbering crosses into cells and `type` reads
 *     `column+line`.
 *
 * Both are CACHE-ONLY (no embedded workbook — `data="none"`), which is the
 * shape the escape hatch can wire today: `raw part add` of the chart part,
 * `raw insert` of its `chart` relationship, then `raw insert` of the drawing
 * paragraph referencing the rId. (A workbook needs a part-level rels part,
 * which `raw part add` refuses; `strict-profile.docx` covers the embedded
 * case.) Every default the spec leaves implicit is written out — both engines
 * smooth a line with no `c:smooth`, and Word rounds a frame with no
 * `c:roundedCorners` — so the LibreOffice round-trip (CORE_FIXTURES) renders
 * what the caches say. Rebuild this with `charts add` once that verb exists.
 */

const root = resolve(import.meta.dir, "../../..");
const out = resolve(root, "tests/fixtures/charts.docx");
const cliEntry = resolve(root, "src/index.ts");

const CHART_RELATIONSHIP_TYPE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const CHART_CONTENT_TYPE =
	"application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

async function cli(...args: string[]): Promise<string> {
	const result = await $`bun ${cliEntry} ${args}`.quiet();
	return result.stdout.toString();
}

const QUARTERS = ["Q1", "Q2", "Q3", "Q4"];

/** A string cache: `<c:strRef>` with its formula and points. */
function strRef(formula: string, values: string[]): string {
	const points = values
		.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`)
		.join("");
	return `<c:strRef><c:f>${formula}</c:f><c:strCache><c:ptCount val="${values.length}"/>${points}</c:strCache></c:strRef>`;
}

/** A number cache; a `null` value is a gap (no `c:pt` for that index). */
function numRef(formula: string, values: (number | null)[]): string {
	const points = values
		.map((value, index) =>
			value === null ? "" : `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`,
		)
		.join("");
	return `<c:numRef><c:f>${formula}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${points}</c:numCache></c:numRef>`;
}

function series(
	index: number,
	column: string,
	name: string,
	values: (number | null)[],
	extra = "",
): string {
	return [
		`<c:ser><c:idx val="${index}"/><c:order val="${index}"/>`,
		`<c:tx>${strRef(`Sheet1!$${column}$1`, [name])}</c:tx>`,
		extra,
		`<c:cat>${strRef("Sheet1!$A$2:$A$5", QUARTERS)}</c:cat>`,
		`<c:val>${numRef(`Sheet1!$${column}$2:$${column}$5`, values)}</c:val>`,
		"</c:ser>",
	].join("");
}

const AXES =
	'<c:catAx><c:axId val="500"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="501"/></c:catAx>' +
	'<c:valAx><c:axId val="501"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="500"/></c:valAx>';

function chartSpace(title: string | null, plot: string): string {
	const titleXml = title
		? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${title}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
		: '<c:autoTitleDeleted val="1"/>';
	return [
		'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
		'<c:date1904 val="0"/><c:roundedCorners val="0"/>',
		`<c:chart>${titleXml}<c:plotArea><c:layout/>${plot}${AXES}</c:plotArea>`,
		'<c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/></c:chart>',
		"</c:chartSpace>",
	].join("");
}

const COLUMN_CHART = chartSpace(
	"Revenue by Quarter",
	[
		'<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>',
		series(0, "B", "West", [120, 135, 150, 170]),
		series(1, "C", "East", [90, 110, null, 140]),
		'<c:axId val="500"/><c:axId val="501"/></c:barChart>',
	].join(""),
);

const COMBO_CHART = chartSpace(
	null,
	[
		'<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>',
		series(0, "B", "Units", [40, 45, 52, 61]),
		'<c:axId val="500"/><c:axId val="501"/></c:barChart>',
		'<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>',
		series(
			1,
			"C",
			"Margin",
			[0.21, 0.24, 0.22, 0.27],
			'<c:marker><c:symbol val="none"/></c:marker>',
		).replace("</c:ser>", '<c:smooth val="0"/></c:ser>'),
		'<c:marker val="1"/><c:axId val="500"/><c:axId val="501"/></c:lineChart>',
	].join(""),
);

/** An inline `<w:drawing>` paragraph pointing at chart relationship `rId`. */
function chartParagraph(rId: string, drawingId: number, name: string): string {
	return [
		'<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">',
		'<wp:extent cx="5486400" cy="3200400"/><wp:effectExtent l="0" t="0" r="0" b="0"/>',
		`<wp:docPr id="${drawingId}" name="${name}"/><wp:cNvGraphicFramePr/>`,
		'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">',
		`<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rId}"/>`,
		"</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>",
	].join("");
}

/** `raw part add` the chart part, then `raw insert` its relationship; returns
 * the minted rId. */
async function addChartPart(partNumber: number, xml: string): Promise<string> {
	await cli(
		"raw",
		"part",
		"add",
		out,
		"--name",
		`word/charts/chart${partNumber}.xml`,
		"--xml",
		xml,
		"--content-type",
		CHART_CONTENT_TYPE,
	);
	const relOutput = await cli(
		"raw",
		"insert",
		out,
		"--xml",
		`<Relationship Type="${CHART_RELATIONSHIP_TYPE}" Target="charts/chart${partNumber}.xml"/>`,
	);
	const rId = relOutput.trim().split("\n")[0];
	if (!rId?.match(/^rId\d+$/))
		throw new Error(`expected a minted rId, got: ${relOutput}`);
	return rId;
}

const REPORT_MARKDOWN = [
	"# Quarterly Report",
	"",
	"Revenue grew every quarter; East has no Q3 figure yet.",
	"",
	"| Units and margin |",
	"| --- |",
	"| placeholder |",
	"",
	"Figures are unaudited.",
].join("\n");
const markdownPath = resolve(
	root,
	"tests/fixtures/setup/.charts-report.tmp.md",
);
await Bun.write(markdownPath, REPORT_MARKDOWN);
await cli("create", out, "--force", "--from", markdownPath);
await Bun.file(markdownPath).delete();

const columnRid = await addChartPart(1, COLUMN_CHART);
await cli(
	"raw",
	"insert",
	out,
	"--after",
	"p1",
	"--xml",
	chartParagraph(columnRid, 101, "Chart 1"),
);

const comboRid = await addChartPart(2, COMBO_CHART);
await cli(
	"raw",
	"replace",
	out,
	"--at",
	"t0:r1c0:p0",
	"--xml",
	chartParagraph(comboRid, 102, "Chart 2"),
);

console.log(`built ${out}`);
