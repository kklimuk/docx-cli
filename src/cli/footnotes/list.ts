import type { NoteKind } from "@core/notes";
import {
	EXIT,
	fail,
	openOrFail,
	respond,
	tryParseArgs,
	writeStdout,
} from "../respond";

function helpFor(kind: NoteKind): string {
	const verb = kind === "footnote" ? "footnotes" : "endnotes";
	const idPrefix = kind === "footnote" ? "fn" : "en";
	return `docx ${verb} list — print existing ${kind}s as JSON

Usage:
  docx ${verb} list FILE [options]

Options:
  -h, --help          Show this help

Output:
  A bare JSON array, one object per ${kind}: { id, text }. Each item's \`id\`
  (e.g. ${idPrefix}0) is the addressable handle you pass as \`--at ${idPrefix}N\` to
  \`docx ${verb} edit\` / \`delete\`. Empty array if the doc has no ${kind}s part.

Examples:
  docx ${verb} list doc.docx | jq '.[] | select(.text | test("citation"))'
`;
}

export async function runListNotes(
	args: string[],
	kind: NoteKind,
): Promise<number> {
	const parsed = await tryParseArgs(
		args,
		{
			help: { type: "boolean", short: "h" },
		},
		helpFor(kind),
	);
	if (typeof parsed === "number") return parsed;

	if (parsed.values.help) {
		await writeStdout(helpFor(kind));
		return EXIT.OK;
	}

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", helpFor(kind));

	const document = await openOrFail(path);
	if (typeof document === "number") return document;

	const notes =
		kind === "footnote" ? document.body.footnotes : document.body.endnotes;
	await respond(notes);
	return EXIT.OK;
}

export async function run(args: string[]): Promise<number> {
	return runListNotes(args, "footnote");
}
