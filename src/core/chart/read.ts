import type { ChartPart } from "../ast/document/charts";
import { resolveRelationshipTarget } from "../ast/document/relationships";
import type { ChartSeries, NativeChartRun } from "../ast/types";
import { alternateContentBranch, isAlternateContent } from "../mc";
import { XmlNode } from "../parser";

/** What a chart part says about itself: the slice of `NativeChartRun` the
 * reader can't get from the `<w:drawing>` alone. */
export type ChartContent = Pick<
	NativeChartRun,
	"chartTypes" | "title" | "series" | "dataSource" | "workbook"
>;

/** Parse one `<c:chartSpace>` part into its chart groups, title, cached series
 * and data source. Read-only: the caches are what Word and LibreOffice draw, so
 * they're what `read` reports; the embedded workbook is only located, never
 * opened. A chart part we can't parse reads as an empty, cache-less chart. */
export function readChartContent(
	part: ChartPart | undefined,
	partName: string,
): ChartContent {
	const chartSpace = part
		? XmlNode.findRoot(part.tree, "c:chartSpace")
		: undefined;
	if (!part || !chartSpace) {
		return { chartTypes: [], series: [], dataSource: "none" };
	}
	const chart = chartSpace.findChild("c:chart");
	const groups = chartGroups(chart?.findChild("c:plotArea"));
	const content: ChartContent = {
		chartTypes: groups.map(chartTypeLabel),
		series: groups.flatMap((group) =>
			group.findChildren("c:ser").map(readSeries),
		),
		...readDataSource(chartSpace, part, partName),
	};
	const title = chart ? readTitle(chart) : undefined;
	if (title) content.title = title;
	return content;
}

/** The plot area's chart-group elements (`<c:barChart>`, `<c:pie3DChart>`, …)
 * in order, looking through any markup-compatibility wrapper. Axes, layout and
 * the data table are siblings and are skipped. */
function chartGroups(plotArea: XmlNode | undefined): XmlNode[] {
	if (!plotArea) return [];
	const groups: XmlNode[] = [];
	visit(plotArea.children);
	return groups;

	function visit(children: XmlNode[]): void {
		for (const child of children) {
			if (isAlternateContent(child)) {
				const branch = alternateContentBranch(child);
				if (branch) visit(branch.children);
				continue;
			}
			if (/^c:\w+Chart$/.test(child.tag)) groups.push(child);
		}
	}
}

/** `c:pie3DChart` → `pie3D`; an upright `c:barChart` → `column` (see
 * `NativeChartRun.chartTypes`). `barDir` is required by the schema; when a
 * producer omits its `val`, the CT_BarDir default is `col`. */
function chartTypeLabel(group: XmlNode): string {
	const name = group.tag.slice("c:".length, -"Chart".length);
	if (name !== "bar" && name !== "bar3D") return name;
	const direction = group.findChild("c:barDir")?.getAttribute("val") ?? "col";
	return direction === "col" ? name.replace("bar", "column") : name;
}

function readSeries(series: XmlNode): ChartSeries {
	const out: ChartSeries = { categories: [], values: [] };
	const name = series.findChild("c:tx");
	if (name) {
		const literal = name.findChild("c:v")?.collectText();
		const cached = readStringPoints(name)[0];
		const text = literal ?? cached;
		if (text) out.name = text;
		const nameRef = formulaOf(name);
		if (nameRef) out.nameRef = nameRef;
	}
	// Scatter and bubble charts carry x/y values instead of categories/values.
	const categories = series.findChild("c:cat") ?? series.findChild("c:xVal");
	if (categories) {
		out.categories = readStringPoints(categories);
		const categoriesRef = formulaOf(categories);
		if (categoriesRef) out.categoriesRef = categoriesRef;
	}
	const values = series.findChild("c:val") ?? series.findChild("c:yVal");
	if (values) {
		out.values = readStringPoints(values).map(parseCachedNumber);
		const valuesRef = formulaOf(values);
		if (valuesRef) out.valuesRef = valuesRef;
		const formatCode = findCache(values)
			?.findChild("c:formatCode")
			?.collectText();
		if (formatCode) out.formatCode = formatCode;
	}
	return out;
}

/** A cached `c:v` as a number; an empty or non-numeric value is a gap. */
function parseCachedNumber(value: string | null): number | null {
	if (value === null || value.trim() === "") return null;
	const number = Number(value);
	return Number.isFinite(number) ? number : null;
}

/** The `c:f` formula of a data source (`c:strRef` / `c:numRef` /
 * `c:multiLvlStrRef`), if it references the workbook. */
function formulaOf(source: XmlNode): string | undefined {
	for (const reference of source.children) {
		const formula = reference.findChild("c:f")?.collectText();
		if (formula) return formula;
	}
	return undefined;
}

const CACHE_TAGS = [
	"c:strCache",
	"c:numCache",
	"c:multiLvlStrCache",
	"c:strLit",
	"c:numLit",
] as const;

/** The point container of a data source: a reference's cache, or a literal
 * (`c:strLit` / `c:numLit`) sitting directly in the source. */
function findCache(source: XmlNode): XmlNode | undefined {
	for (const container of [source, ...source.children]) {
		for (const tag of CACHE_TAGS) {
			const cache = container.findChild(tag);
			if (cache) return cache;
		}
	}
	return undefined;
}

/** A cache's points as strings, indexed by `c:pt/@idx` and padded to
 * `c:ptCount` — a missing point (a gap in the data) is `null`. A multi-level
 * category cache reads its first `c:lvl`: the innermost labels, the ones each
 * point sits on. */
function readStringPoints(source: XmlNode): (string | null)[] {
	let cache = findCache(source);
	if (cache?.tag === "c:multiLvlStrCache")
		cache = cache.findChild("c:lvl") ?? cache;
	if (!cache) return [];
	const declared = Number(
		cache.findChild("c:ptCount")?.getAttribute("val") ?? "0",
	);
	const points: (string | null)[] = [];
	for (const point of cache.findChildren("c:pt")) {
		const index = Number(point.getAttribute("idx"));
		if (!isPointIndex(index)) continue;
		points[index] = point.findChild("c:v")?.collectText() ?? "";
	}
	const length = Math.max(points.length, isPointIndex(declared) ? declared : 0);
	return Array.from({ length }, (_, index) => points[index] ?? null);
}

/** A chart can't plot more points than a worksheet has rows (1,048,576), so
 * a larger `idx`/`ptCount` is corrupt input — skip it rather than allocate. */
const MAX_POINTS = 1_048_576;

function isPointIndex(value: number): boolean {
	return Number.isInteger(value) && value >= 0 && value < MAX_POINTS;
}

/** The title's text: rich-text runs (`c:tx/c:rich`), one string per
 * paragraph joined by a space, or a formula-backed title's cached string.
 * An auto title (no `c:tx`) or a deleted one reads as none. */
function readTitle(chart: XmlNode): string | undefined {
	const text = chart.findChild("c:title")?.findChild("c:tx");
	if (!text) return undefined;
	const rich = text.findChild("c:rich");
	if (rich) {
		const paragraphs = rich
			.findChildren("a:p")
			.map((paragraph) => collectDrawingText(paragraph).trim())
			.filter((paragraph) => paragraph.length > 0);
		return paragraphs.length > 0 ? paragraphs.join(" ") : undefined;
	}
	const cached = readStringPoints(text)[0];
	return cached ? cached : undefined;
}

/** Text of a DrawingML paragraph: every `a:t`, with `a:br` as a space. */
function collectDrawingText(paragraph: XmlNode): string {
	let out = "";
	for (const child of paragraph.children) {
		if (child.tag === "a:r" || child.tag === "a:fld")
			out += child.findChild("a:t")?.collectText() ?? "";
		else if (child.tag === "a:br") out += " ";
	}
	return out;
}

/** Where Word's Edit Data reads from: `<c:externalData r:id>` → the chart
 * part's rels → an internal workbook part (`embedded`) or a
 * `TargetMode="External"` link (`external`); no `externalData` is a
 * cache-only chart (`none`). */
function readDataSource(
	chartSpace: XmlNode,
	part: ChartPart,
	partName: string,
): Pick<ChartContent, "dataSource" | "workbook"> {
	const relationshipId = chartSpace
		.findChild("c:externalData")
		?.getAttribute("r:id");
	if (!relationshipId) return { dataSource: "none" };
	const relationship = part.relationships.findByRid(relationshipId);
	const target = relationship?.getAttribute("Target");
	if (!relationship || !target) return { dataSource: "none" };
	if (relationship.getAttribute("TargetMode") === "External")
		return { dataSource: "external" };
	return {
		dataSource: "embedded",
		workbook: resolveRelationshipTarget(partName, target),
	};
}
