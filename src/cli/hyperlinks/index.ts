import { fail, writeStdout } from "../respond";

type CommandFn = (args: string[]) => Promise<number>;

const SUBCOMMANDS: Record<string, () => Promise<{ run: CommandFn }>> = {
	add: () => import("./add"),
	delete: () => import("./delete"),
	list: () => import("./list"),
	replace: () => import("./replace"),
};

const HELP = `docx hyperlinks — manage hyperlinks

Usage:
  docx hyperlinks <verb> FILE [options]

Verbs:
  add      Wrap an existing span in a hyperlink
  delete   Unwrap a hyperlink (keep the text)
  list     Print hyperlink manifest as JSON
  replace  Change a hyperlink's URL

Run "docx hyperlinks <verb> --help" for verb-specific help.
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
			`Unknown hyperlinks subcommand: ${verb}`,
			'Run "docx hyperlinks --help".',
		);
	}
	const module_ = await loader();
	return module_.run(args.slice(1));
}
