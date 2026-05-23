import { fail, writeStdout } from "../respond";

type CommandFn = (args: string[]) => Promise<number>;

const SUBCOMMANDS: Record<string, () => Promise<{ run: CommandFn }>> = {
	add: () => import("./add"),
	delete: () => import("./delete"),
	edit: () => import("./edit"),
	list: () => import("./list"),
};

const HELP = `docx footnotes — author footnotes

Usage:
  docx footnotes <verb> FILE [options]

Verbs:
  add      Insert a footnote reference anchored to a paragraph offset
  edit     Replace an existing footnote's body text
  list     Print existing footnotes as JSON
  delete   Remove a footnote body and every reference to it

Run "docx footnotes <verb> --help" for verb-specific help.
`;

export async function run(args: string[]): Promise<number> {
	const verb = args[0];
	if (!verb || verb === "--help" || verb === "-h" || verb === "help") {
		await writeStdout(HELP);
		return verb ? 0 : 2;
	}
	const loader = SUBCOMMANDS[verb];
	if (!loader) {
		return fail(
			"USAGE",
			`Unknown footnotes subcommand: ${verb}`,
			'Run "docx footnotes --help".',
		);
	}
	const module_ = await loader();
	return module_.run(args.slice(1));
}
