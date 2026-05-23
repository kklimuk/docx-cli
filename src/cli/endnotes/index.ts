import { runAddNote } from "../footnotes/add";
import { runDeleteNote } from "../footnotes/delete";
import { runEditNote } from "../footnotes/edit";
import { runListNotes } from "../footnotes/list";
import { fail, writeStdout } from "../respond";

type CommandFn = (args: string[]) => Promise<number>;

// Footnotes and endnotes share the same mechanics — the only difference is
// the part name and a handful of tag/style names, all parameterized in
// `cli/footnotes/helpers.tsx` as `NoteKind`. So this dispatcher re-uses the
// footnote command implementations and just passes kind="endnote".
const SUBCOMMANDS: Record<string, CommandFn> = {
	add: (args) => runAddNote(args, "endnote"),
	delete: (args) => runDeleteNote(args, "endnote"),
	edit: (args) => runEditNote(args, "endnote"),
	list: (args) => runListNotes(args, "endnote"),
};

const HELP = `docx endnotes — author endnotes

Usage:
  docx endnotes <verb> FILE [options]

Verbs:
  add      Insert an endnote reference anchored to a paragraph offset
  edit     Replace an existing endnote's body text
  list     Print existing endnotes as JSON
  delete   Remove an endnote body and every reference to it

Run "docx endnotes <verb> --help" for verb-specific help.
`;

export async function run(args: string[]): Promise<number> {
	const verb = args[0];
	if (!verb || verb === "--help" || verb === "-h" || verb === "help") {
		await writeStdout(HELP);
		return verb ? 0 : 2;
	}
	const handler = SUBCOMMANDS[verb];
	if (!handler) {
		return fail(
			"USAGE",
			`Unknown endnotes subcommand: ${verb}`,
			'Run "docx endnotes --help".',
		);
	}
	return handler(args.slice(1));
}
