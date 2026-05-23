import { parseArgs } from "util";
import { EXIT, fail, openOrFail, respond, writeStdout } from "../respond";
import type { NoteKind } from "./helpers";

function helpFor(kind: NoteKind): string {
	const verb = kind === "footnote" ? "footnotes" : "endnotes";
	return `docx ${verb} list — print existing ${kind}s as JSON

Usage:
  docx ${verb} list FILE [options]

Options:
  -h, --help          Show this help

Examples:
  docx ${verb} list doc.docx | jq '.[] | select(.text | test("citation"))'
`;
}

export async function runListNotes(
	args: string[],
	kind: NoteKind,
): Promise<number> {
	let parsed: ReturnType<typeof parseArgs>;
	try {
		parsed = parseArgs({
			args,
			allowPositionals: true,
			options: {
				help: { type: "boolean", short: "h" },
			},
		});
	} catch (parseError) {
		const message =
			parseError instanceof Error ? parseError.message : String(parseError);
		return fail("USAGE", message, helpFor(kind));
	}

	if (parsed.values.help) {
		await writeStdout(helpFor(kind));
		return EXIT.OK;
	}

	const path = parsed.positionals[0];
	if (!path) return fail("USAGE", "Missing FILE argument", helpFor(kind));

	const view = await openOrFail(path);
	if (typeof view === "number") return view;

	const notes = kind === "footnote" ? view.doc.footnotes : view.doc.endnotes;
	await respond(notes);
	return EXIT.OK;
}

export async function run(args: string[]): Promise<number> {
	return runListNotes(args, "footnote");
}
