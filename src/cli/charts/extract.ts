import { Charts, chartToCsv } from "@core/chart";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx charts extract — print one chart's data

Usage:
  docx charts extract FILE --at chartN [--json]

Required:
  --at chartN       The chart (e.g. chart0; see \`docx charts list\`)

Optional:
  --json            Print the full series objects (name, categories, values,
                    workbook refs, formatCode) instead of CSV
  -h, --help        Show this help

Output:
  CSV (RFC 4180) by default: a header row "category,<series 1>,<series 2>,…",
  then one row per category. An empty cell is a gap in the data, not 0.
  The values are the chart's CACHE — what Word draws on the page. Redirect to
  save it: docx charts extract report.docx --at chart0 > chart0.csv
  Errors print {code, error, hint?} with a nonzero exit.

Examples:
  docx charts extract report.docx --at chart0
  docx charts extract report.docx --at chart1 --json | jq '.series[0].values'
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			at: { type: "string" },
			json: { type: "boolean" },
			help: { type: "boolean", short: "h" },
		},
		HELP,
	);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const chartId = parsed.values.at as string | undefined;
	if (!chartId) return fail("USAGE", "Missing --at chartN", HELP);

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const chart = new Charts(document).get(chartId);
	if (!chart) {
		return fail(
			"CHART_NOT_FOUND",
			`Chart not found: ${chartId}`,
			`Run "docx charts list ${path}" to see the chart ids.`,
		);
	}

	if (parsed.values.json) {
		await respond({ id: chart.id, title: chart.title, series: chart.series });
		return EXIT.OK;
	}
	await writeStdout(chartToCsv(chart));
	return EXIT.OK;
}
