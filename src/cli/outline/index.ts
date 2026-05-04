import { parseArgs } from "util";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";
import { buildOutline } from "./build";

const HELP = `docx outline — list headings as a hierarchical tree

Usage:
  docx outline FILE [options]

Options:
  --style-prefix S  paragraph-style prefix that marks a heading (default: "Heading")
  -h, --help        show this help

Walks top-level paragraphs whose style starts with the prefix and parses the
trailing number as the heading level (e.g. "Heading1" → 1, "Heading 2" → 2).
Paragraphs nested inside table cells are skipped — outlines reflect the
document's structural skeleton, not embedded labels. Lower levels nest under
higher ones; missing intermediate levels nest directly (H1 → H3 is fine).

Output is JSON: an array of entries, each shaped like
  { id, locator, level, style, text, children }.

Examples:
  docx outline doc.docx
  docx outline doc.docx --style-prefix "Section"
  docx outline doc.docx | jq '.outline[].text'
`;

export async function run(args: string[]): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				"style-prefix": { type: "string" },
				help: { type: "boolean", short: "h" },
			},
		});
	} catch (parseError) {
		const message =
			parseError instanceof Error ? parseError.message : String(parseError);
		return fail("USAGE", message, HELP);
	}

	if (parsed.values.help) {
		await writeStdout(HELP);
		return EXIT.OK;
	}

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", HELP);

	const stylePrefix =
		(parsed.values["style-prefix"] as string | undefined) ?? "Heading";
	if (stylePrefix.length === 0) {
		return fail("USAGE", "--style-prefix cannot be empty");
	}

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const outline = buildOutline(view.doc, { stylePrefix });
	await respond({
		ok: true,
		operation: "outline",
		path,
		stylePrefix,
		outline,
	});
	return EXIT.OK;
}
