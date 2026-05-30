import { Comments } from "@core";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";

const HELP = `docx comments list — print existing comments as JSON

Usage:
  docx comments list FILE [options]

Options:
  --include-resolved   Include resolved comments (default excludes them)
  --thread cN          Print only the thread rooted at the given comment id
  -h, --help           Show this help

Examples:
  docx comments list doc.docx
  docx comments list doc.docx --include-resolved | jq '.[] | select(.author == "Jane")'
`;

export async function run(args: string[]): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			"include-resolved": { type: "boolean" },
			thread: { type: "string" },
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

	const comments = new Comments(document).list({
		includeResolved: Boolean(parsed.values["include-resolved"]),
		thread: parsed.values.thread as string | undefined,
	});

	await respond(comments);
	return EXIT.OK;
}
