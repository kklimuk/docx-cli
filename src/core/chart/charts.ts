import type { Document } from "../ast/document";
import { iterateBlocks } from "../ast/document/body";
import type { ChartSeries, NativeChartRun } from "../ast/types";

/** The chart lens: read-only access to the document's native charts by their
 * `chartN` handle. Stateless — it reads the AST the reader built (every chart
 * run carries its parsed caches), so it never re-opens a chart part. */
export class Charts {
	constructor(private readonly document: Document) {}

	/** Every native chart in document order (the order `chartN` ids are minted),
	 * with the block that holds it. */
	list(): ChartListing[] {
		return [...this.charts()].map(({ run, at }) => toListing(run, at));
	}

	/** The chart run for `chartN`, or undefined when no chart has that id. */
	get(chartId: string): NativeChartRun | undefined {
		for (const { run } of this.charts()) if (run.id === chartId) return run;
		return undefined;
	}

	private *charts(): IterableIterator<{ run: NativeChartRun; at: string }> {
		for (const block of iterateBlocks(this.document.body.blocks)) {
			if (block.type !== "paragraph") continue;
			for (const run of block.runs) {
				if (run.type === "chart" && run.kind === "chart")
					yield { run, at: block.id };
			}
		}
	}
}

function toListing(run: NativeChartRun, at: string): ChartListing {
	return {
		id: run.id,
		at,
		type: chartTypeName(run),
		...(run.title ? { title: run.title } : {}),
		series: run.series.map((series, index) => seriesName(series, index)),
		points: chartPointCount(run),
		data: run.dataSource,
		...(run.workbook ? { workbook: run.workbook } : {}),
	};
}

/** The chart's type as one token: its chart groups joined by `+`
 * (`column+line` for a combo), or `unknown` when the part didn't resolve. */
export function chartTypeName(run: NativeChartRun): string {
	return run.chartTypes.length > 0 ? run.chartTypes.join("+") : "unknown";
}

/** How many points (categories) the chart plots — the longest series. */
export function chartPointCount(run: NativeChartRun): number {
	let count = 0;
	for (const series of run.series) {
		count = Math.max(count, series.categories.length, series.values.length);
	}
	return count;
}

/** A series' display name, falling back to Word's own `Series N` label. */
export function seriesName(series: ChartSeries, index: number): string {
	return series.name ?? `Series ${index + 1}`;
}

/** One `charts list` item; `id` is the `--at` handle, `at` the paragraph that
 * holds the chart. */
type ChartListing = {
	id: string;
	at: string;
	type: string;
	title?: string;
	series: string[];
	points: number;
	data: NativeChartRun["dataSource"];
	workbook?: string;
};
