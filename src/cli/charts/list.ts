import { Charts } from "@core/chart";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx charts list — print the chart manifest as JSON

Usage:
  docx charts list FILE [options]

Options:
  -h, --help        Show this help

Output:
  A bare JSON array, one object per native chart in document order:
  { id, at, type, title?, series, points, data, workbook? }.
    id        chartN — the --at handle for \`charts extract\`
    at        the paragraph holding the chart (pN, t0:r0c0:p0, tbx0:p0, …)
    type      chart type; a combo chart joins its types with "+" (column+line)
    series    series names
    points    number of categories plotted
    data      where Word's Edit Data reads from: embedded | external | none
    workbook  the embedded workbook's part name, when data is "embedded"
  Errors print {code, error, hint?} with a nonzero exit.

Examples:
  docx charts list report.docx | jq -c '.[] | {id, type, title}'
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
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

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	await respond(new Charts(document).list());
	return EXIT.OK;
}
