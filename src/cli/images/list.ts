import { flattenImageRuns } from "@core";
import { Images } from "@core/image";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx images list — print image manifest as JSON

Usage:
  docx images list FILE [options]

Options:
  -h, --help        Show this help

Output:
  A bare JSON array of image objects: { id, contentType, hash, ... }. Each
  item's "id" (e.g. img0) is its addressable handle — pass it to
  \`images extract/replace/delete --at\`. Errors print {code, error, hint?}
  with a nonzero exit.

Examples:
  docx images list doc.docx | jq -c '.[] | {id, contentType, hash}'
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

	await new Images(document).enrichHashes();

	await respond(flattenImageRuns(document.body.blocks));
	return EXIT.OK;
}
