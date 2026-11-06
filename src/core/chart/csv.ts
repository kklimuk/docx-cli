import type { NativeChartRun } from "../ast/types";
import { chartPointCount, seriesName } from "./charts";

/** The chart's cached data as RFC 4180 CSV: a header row
 * (`category,<series 1>,<series 2>…`), then one row per point with its
 * category label and each series' value. A gap (a point missing from a
 * cache) is an empty cell, never `0`. Series that label their points
 * differently share the first label any series gives that point. Rows end in
 * LF (the RFC's CRLF is a nuisance in a terminal pipeline; RFC parsers accept
 * either). */
export function chartToCsv(run: NativeChartRun): string {
	const header = [
		"category",
		...run.series.map((series, index) => seriesName(series, index)),
	];
	const rows = [header];
	for (let index = 0; index < chartPointCount(run); index++) {
		const category =
			run.series.find((series) => series.categories[index] != null)?.categories[
				index
			] ?? "";
		rows.push([
			category,
			...run.series.map((series) => {
				const value = series.values[index];
				return value == null ? "" : String(value);
			}),
		]);
	}
	return rows.map((row) => `${row.map(csvField).join(",")}\n`).join("");
}

/** Quote a field when it holds a comma, quote or line break (RFC 4180 §2.6),
 * doubling embedded quotes. */
function csvField(value: string): string {
	if (!/[",\r\n]/.test(value)) return value;
	return `"${value.replace(/"/g, '""')}"`;
}
