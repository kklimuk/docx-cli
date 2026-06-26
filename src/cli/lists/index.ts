import { fail, writeStdout } from "../respond";

type CommandFn = (args: string[]) => Promise<number>;

const SUBCOMMANDS: Record<string, () => Promise<{ run: CommandFn }>> = {
	set: () => import("./set"),
};

const HELP = `docx lists — control list numbering (start value, glyph format, restart/continue)

Usage:
  docx lists set FILE --at pN [options]

Verbs:
  set   Renumber a numbered list — set its start value or glyph format, or make it
        restart vs. continue the previous list
        (--at pN [--start N] [--format FMT] [--restart] [--continue])

Lists are otherwise handled by the standard verbs:
  create a list   docx insert FILE --after pN --list "first,second"  (or markdown "1. …")
  edit an item    docx edit FILE --at pN --text "..."
  inspect         docx read FILE --ast   (list.numId / level / ordered / start / format)

Run "docx lists set --help" for option detail.
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
			`Unknown lists subcommand: ${verb}`,
			'Run "docx lists --help".',
		);
	}
	const module_ = await loader();
	return module_.run(args.slice(1));
}
